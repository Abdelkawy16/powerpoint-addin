/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
    
    // Add event handler for document selection changes
    Office.context.document.addHandlerAsync(
      Office.EventType.DocumentSelectionChanged,
      selectionChangedHandler,
      function(result) {
        if (result.status === Office.AsyncResultStatus.Failed) {
          console.error("Failed to add selection change handler: " + result.error.message);
        }
      }
    );
  }
});

// Global variables to track selection state
let lastSelectedText = "";
let selectionChangeTimeout = null;

// Handler for selection changed events
function selectionChangedHandler(eventArgs) {
  // Debounce the selection change events to prevent too many API calls
  if (selectionChangeTimeout) {
    clearTimeout(selectionChangeTimeout);
  }
  
  selectionChangeTimeout = setTimeout(() => {
    // Get the current selection when the selection changes
    detectSelectedElement();
  }, 300); // 300ms debounce
}

// Function to detect the currently selected element
function detectSelectedElement() {
  try {
    document.getElementById('item-subject').textContent = "Detecting selected element...";
    
    // First try to detect shapes directly
    detectShapesDirectly();
    
    // Also try to get text selection using the common Office API
    Office.context.document.getSelectedDataAsync(
      Office.CoercionType.Text,
      { valueFormat: "unformatted" },
      function(asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
          const selectedText = asyncResult.value || "";
          
          // Only process if the selection has actually changed and there is text
          if (selectedText !== lastSelectedText && selectedText.length > 0) {
            lastSelectedText = selectedText;
            document.getElementById('item-subject').textContent = 
              `Selected text: "${selectedText}"\n\n`;
            
            // Continue with shape detection to get more details
            getSelectedShapeDetails(selectedText);
          }
        }
      }
    );
  } catch (error) {
    console.error("Selection detection error:", error);
    document.getElementById('item-subject').textContent = 'Error: ' + (error.message || error);
  }
}

// Function to detect shapes directly
function detectShapesDirectly() {
  PowerPoint.run(async (context) => {
    try {
      // Get the selected slides
      const selectedSlides = context.presentation.getSelectedSlides();
      selectedSlides.load("items");
      await context.sync();
      
      // Check if there are any selected slides
      if (!selectedSlides.items || selectedSlides.items.length === 0) {
        document.getElementById('item-subject').textContent = "No slide is currently selected.";
        return;
      }
      
      // Get the first selected slide
      const slide = selectedSlides.items[0];
      
      // Try multiple approaches to detect the selected shape
      
      // Approach 1: Try to get the selection directly (newer PowerPoint versions)
      try {
        const selection = context.presentation.getSelection();
        selection.load("shapes");
        await context.sync();
        
        if (selection.shapes && selection.shapes.items && selection.shapes.items.length > 0) {
          // We found directly selected shapes!
          const selectedShapes = selection.shapes.items;
          
          // Load detailed properties for each selected shape
          for (let i = 0; i < selectedShapes.length; i++) {
            selectedShapes[i].load("name,id,type,text,left,top,width,height,zIndex");
          }
          await context.sync();
          
          // Display information about the selected shapes
          let message = `Found ${selectedShapes.length} selected element(s):\n\n`;
          
          selectedShapes.forEach((shape, index) => {
            message += getDetailedShapeInfo(shape, index);
          });
          
          document.getElementById('item-subject').textContent = message;
          return; // Exit early as we found what we needed
        }
      } catch (selectionError) {
        console.log("Direct selection API not supported:", selectionError);
        // Continue to other methods
      }
      
      // Approach 2: Try to get active view and selection (alternative method)
      try {
        const view = context.presentation.getActiveView();
        view.load("selection");
        await context.sync();
        
        if (view.selection) {
          view.selection.load("shapes");
          await context.sync();
          
          if (view.selection.shapes && view.selection.shapes.items && view.selection.shapes.items.length > 0) {
            const selectedShapes = view.selection.shapes.items;
            
            // Load detailed properties for each selected shape
            for (let i = 0; i < selectedShapes.length; i++) {
              selectedShapes[i].load("name,id,type,text,left,top,width,height,zIndex");
            }
            await context.sync();
            
            // Display information about the selected shapes
            let message = `Found ${selectedShapes.length} selected element(s) via active view:\n\n`;
            
            selectedShapes.forEach((shape, index) => {
              message += getDetailedShapeInfo(shape, index);
            });
            
            document.getElementById('item-subject').textContent = message;
            return; // Exit early as we found what we needed
          }
        }
      } catch (viewError) {
        console.log("Active view selection API not supported:", viewError);
        // Continue to other methods
      }
      
      // Approach 3: Check all shapes on the slide for isSelected property
      slide.load("shapes");
      await context.sync();
      
      if (!slide.shapes || !slide.shapes.items || slide.shapes.items.length === 0) {
        document.getElementById('item-subject').textContent = "No elements found on the current slide.";
        return;
      }
      
      // Load all shapes with their properties
      const shapes = slide.shapes.items;
      let foundSelectedShape = false;
      
      // Try to find selected shapes by checking isSelected property
      for (let i = 0; i < shapes.length; i++) {
        const shape = shapes[i];
        // Load all relevant properties including isSelected if available
        shape.load("name,id,type,text,left,top,width,height,zIndex,isSelected");
      }
      
      await context.sync();
      
      // Check if any shapes are marked as selected
      const selectedShapes = shapes.filter(shape => shape.isSelected);
      
      if (selectedShapes.length > 0) {
        // We found selected shapes!
        let message = `Found ${selectedShapes.length} selected element(s):\n\n`;
        
        selectedShapes.forEach((shape, index) => {
          message += getDetailedShapeInfo(shape, index);
        });
        
        document.getElementById('item-subject').textContent = message;
        foundSelectedShape = true;
      }
      
      // Approach 4: Try to detect the most recently clicked shape based on z-order
      // This is a heuristic approach that might help in some cases
      if (!foundSelectedShape) {
        try {
          // Sort shapes by z-index (higher z-index is typically the one on top that was clicked)
          const sortedShapes = [...shapes].sort((a, b) => b.zIndex - a.zIndex);
          
          if (sortedShapes.length > 0) {
            // Take the top-most shape as a best guess
            const topShape = sortedShapes[0];
            topShape.load("name,id,type,text,left,top,width,height,zIndex");
            await context.sync();
            
            let message = `Best guess of selected element (top-most element):\n\n`;
            message += getDetailedShapeInfo(topShape, 0);
            message += `Note: This is a best guess based on the top-most element on the slide.\n`;
            message += `For more accurate results, try clicking directly on an element before pressing 'Get Info'.\n\n`;
            
            document.getElementById('item-subject').textContent = message;
            foundSelectedShape = true;
          }
        } catch (zOrderError) {
          console.log("Z-order detection failed:", zOrderError);
          // Continue to fallback
        }
      }
      
      // If no selected shape was found, show all shapes with their details
      if (!foundSelectedShape) {
        // Group shapes by type for better organization
        const shapesByType = {};
        
        shapes.forEach((shape, index) => {
          const type = shape.type || "Unknown";
          if (!shapesByType[type]) {
            shapesByType[type] = [];
          }
          shapesByType[type].push({ shape, index });
        });
        
        let message = `Could not identify which element is selected.\n\n`;
        message += `All elements on current slide (${shapes.length} total):\n\n`;
        
        // Display shapes grouped by type
        for (const type in shapesByType) {
          message += `== ${type} Elements (${shapesByType[type].length}) ==\n\n`;
          
          shapesByType[type].forEach(({ shape, index }) => {
            message += `${index + 1}. ${shape.name || "Unnamed"} (ID: ${shape.id})\n`;
            message += `   Position: (${shape.left}, ${shape.top})\n`;
            message += `   Size: ${shape.width} x ${shape.height}\n`;
            message += `   Z-Index: ${shape.zIndex}\n\n`;
          });
        }
        
        message += "\nTips for selecting elements:\n";
        message += "1. Click directly on the element you want to examine\n";
        message += "2. Press the 'Get Info' button immediately after selecting\n";
        message += "3. For text elements, try selecting some text within the element";
        
        document.getElementById('item-subject').textContent = message;
      }
    } catch (error) {
      console.error("Shape detection error:", error);
      document.getElementById('item-subject').textContent = 'Error detecting shapes: ' + (error.message || error);
    }
  });
}

// Helper function to get detailed information about a shape
function getDetailedShapeInfo(shape, index) {
  let details = `${index + 1}. ${shape.name || "Unnamed"} (ID: ${shape.id})\n`;
  details += `   Type: ${shape.type}\n`;
  
  // Add specific details based on shape type
  switch (shape.type) {
    case "Picture":
      details += `   Element Type: Image/Picture\n`;
      break;
    case "Group":
      details += `   Element Type: Group (contains multiple elements)\n`;
      break;
    case "Table":
      details += `   Element Type: Table\n`;
      break;
    case "Chart":
      details += `   Element Type: Chart\n`;
      break;
    case "SmartArt":
      details += `   Element Type: SmartArt\n`;
      break;
  }
  
  details += `   Position: (${shape.left}, ${shape.top})\n`;
  details += `   Size: ${shape.width} x ${shape.height}\n`;
  details += `   Z-Index: ${shape.zIndex}\n`;
  
  if (shape.text) {
    // Truncate text if it's too long
    const maxTextLength = 100;
    const displayText = shape.text.length > maxTextLength ? 
      shape.text.substring(0, maxTextLength) + "..." : 
      shape.text;
    details += `   Text: "${displayText}"\n`;
  }
  
  details += "\n";
  return details;
}

export async function run() {
  detectSelectedElement();
}

// Function to get details about the selected shape or all shapes if none is selected
async function getSelectedShapeDetails(selectedText = null) {
  try {
    await PowerPoint.run(async (context) => {
      try {
        // Get the selected slides
        const selectedSlides = context.presentation.getSelectedSlides();
        selectedSlides.load("items");
        await context.sync();
        
        // Check if there are any selected slides
        if (!selectedSlides.items || selectedSlides.items.length === 0) {
          document.getElementById('item-subject').textContent += "\n\nNo slide is currently selected.";
          return;
        }
        
        // Get the first selected slide
        const slide = selectedSlides.items[0];
        
        // Load the shapes collection
        slide.load("shapes");
        await context.sync();
        
        if (slide.shapes) {
          slide.shapes.load("items");
          await context.sync();
          
          let message = selectedText ? 
            document.getElementById('item-subject').textContent : 
            "Detecting selected shape... ";
          
          // Check if we have shapes on the slide
          if (slide.shapes.items && slide.shapes.items.length > 0) {
            const shapes = slide.shapes.items;
            
            // If we have selected text, try to find which shape contains it
            let foundSelectedShape = false;
            
            // First attempt: Use selected text to identify the shape
            if (selectedText) {
              message += "Attempting to identify the shape containing this text...\n\n";
              
              // Load more properties for each shape
              for (let i = 0; i < shapes.length; i++) {
                const shape = shapes[i];
                shape.load("name,id,type,text,left,top,width,height,zIndex,isSelected");
                await context.sync();
                
                // Check if this shape contains text and if it matches our selected text
                if (shape.text && shape.text.indexOf(selectedText) !== -1) {
                  message += `Found shape containing selected text:\n\n`;
                  message += getShapeDetailsString(shape, i);
                  foundSelectedShape = true;
                  break;
                }
                
                // Some PowerPoint versions support isSelected property
                if (shape.isSelected) {
                  message += `Found selected shape:\n\n`;
                  message += getShapeDetailsString(shape, i);
                  foundSelectedShape = true;
                  break;
                }
              }
              
              if (!foundSelectedShape) {
                message += "Could not identify which specific shape contains the selected text.\n\n";
              }
            }
            
            // Second attempt: Try to detect selected shape using alternative methods
            if (!foundSelectedShape) {
              // Try to get the active selection via the API (some PowerPoint versions support this)
              try {
                const selection = context.presentation.getSelection();
                selection.load("shapes");
                await context.sync();
                
                if (selection.shapes && selection.shapes.items && selection.shapes.items.length > 0) {
                  message += `Found selected shape via selection API:\n\n`;
                  const selectedShape = selection.shapes.items[0];
                  selectedShape.load("name,id,type,text,left,top,width,height,zIndex");
                  await context.sync();
                  
                  message += getShapeDetailsString(selectedShape, 0);
                  foundSelectedShape = true;
                }
              } catch (selectionError) {
                console.log("Selection API not supported or other error:", selectionError);
                // Continue to fallback method
              }
            }
            
            // If we still didn't find a specific shape, show all shapes
            if (!foundSelectedShape) {
              message += `All elements on current slide (${shapes.length} total):\n\n`;
              
              // Load more properties for each shape
              for (let i = 0; i < shapes.length; i++) {
                const shape = shapes[i];
                shape.load("name,id,type,left,top,width,height,zIndex");
                await context.sync();
                
                message += `${i + 1}. ${shape.name || "Unnamed"} (ID: ${shape.id})\n`;
                message += `   Type: ${shape.type}, Position: (${shape.left}, ${shape.top})\n`;
                message += `   Size: ${shape.width} x ${shape.height}, Z-Index: ${shape.zIndex}\n\n`;
              }
              
              message += "\nTip: Select text within a shape to help identify it more precisely.";
            }
          } else {
            message += "No elements found on the current slide.";
          }
          
          // Update the display
          document.getElementById('item-subject').textContent = message;
        } else {
          document.getElementById('item-subject').textContent += "\n\nCould not access shapes on the current slide.";
        }
      } catch (innerError) {
        console.error("Inner error:", innerError);
        document.getElementById('item-subject').textContent += `\n\nError processing slide: ${innerError.message || innerError}`;
      }
    });
  } catch (error) {
    console.error("Shape detection error:", error);
    document.getElementById('item-subject').textContent += '\n\nError detecting shapes: ' + (error.message || error);
  }
}

// Helper function to format shape details as a string
function getShapeDetailsString(shape, index) {
  return getDetailedShapeInfo(shape, index);
}
