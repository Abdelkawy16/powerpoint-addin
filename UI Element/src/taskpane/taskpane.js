/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

// Global variables to track selection and mouse state
let lastSelectedText = "";
let selectionChangeTimeout = null;
let lastMousePosition = { x: 0, y: 0 };
let lastClickedPosition = { x: 0, y: 0 };
let hasMousePositionData = false;

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
    
    // Set up mouse position tracking in PowerPoint
    setupMouseTracking();
  }
});

// Set up mouse position tracking
function setupMouseTracking() {
  // Add message listener for mouse position data from PowerPoint
  Office.context.document.addHandlerAsync(
    "documentSelectionChanged", 
    function() {
      // This event fires when selection changes, which often happens after a mouse click
      // We'll capture the last mouse position at this time
      try {
        Office.context.document.getSelectedDataAsync(
          Office.CoercionType.SlideRange,
          function(asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
              // Record that we clicked somewhere
              lastClickedPosition = lastMousePosition;
              hasMousePositionData = true;
            }
          }
        );
      } catch (error) {
        console.log("Error in mouse tracking:", error);
      }
    }
  );
  
  // Add a custom button to capture current position
  const capturePositionButton = document.createElement("button");
  capturePositionButton.textContent = "Capture Current Position";
  capturePositionButton.onclick = function() {
    lastClickedPosition = lastMousePosition;
    hasMousePositionData = true;
    detectSelectedElement();
  };
  
  // Insert the button before the main run button
  const runButton = document.getElementById("run");
  if (runButton && runButton.parentNode) {
    runButton.parentNode.insertBefore(capturePositionButton, runButton);
  }
}

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
    
    // Create a selection info object to store all detected information
    const selectionInfo = {
      text: null,
      shapes: [],
      pictures: [],
      detectionMethod: null
    };
    
    // Get all selection information in parallel
    Promise.all([
      detectSelectionText(selectionInfo),
      detectPicturesSpecifically(selectionInfo),
      detectShapesDirectly(selectionInfo)
    ]).then(() => {
      // Once all detection methods have completed, display the results
      displaySelectionInfo(selectionInfo);
    }).catch(error => {
      console.error("Selection detection error:", error);
      document.getElementById('item-subject').textContent = 'Error: ' + (error.message || error);
    });
    
  } catch (error) {
    console.error("Selection detection error:", error);
    document.getElementById('item-subject').textContent = 'Error: ' + (error.message || error);
  }
}

// Function to detect selected text
function detectSelectionText(selectionInfo) {
  return new Promise((resolve) => {
    Office.context.document.getSelectedDataAsync(
      Office.CoercionType.Text,
      { valueFormat: "unformatted" },
      function(asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
          const selectedText = asyncResult.value || "";
          
          if (selectedText.length > 0) {
            selectionInfo.text = selectedText;
            selectionInfo.detectionMethod = "text";
            lastSelectedText = selectedText;
          }
        }
        resolve();
      }
    );
  });
}

// Function to specifically detect pictures
function detectPicturesSpecifically(selectionInfo) {
  return new Promise((resolve) => {
    PowerPoint.run(async (context) => {
      try {
        // Get the selected slides
        const selectedSlides = context.presentation.getSelectedSlides();
        selectedSlides.load("items");
        await context.sync();
        
        // Check if there are any selected slides
        if (!selectedSlides.items || selectedSlides.items.length === 0) {
          resolve();
          return;
        }
        
        // Get the first selected slide
        const slide = selectedSlides.items[0];
        
        // Try to get pictures on the slide
        slide.load("shapes");
        await context.sync();
        
        if (!slide.shapes || !slide.shapes.items || slide.shapes.items.length === 0) {
          resolve();
          return;
        }
        
        // Get all shapes and filter for pictures
        const shapes = slide.shapes.items;
        
        // Load type property for all shapes
        for (let i = 0; i < shapes.length; i++) {
          shapes[i].load("type");
        }
        await context.sync();
        
        // Filter for pictures
        const pictures = shapes.filter(shape => shape.type === "Picture");
        
        if (pictures.length === 0) {
          resolve();
          return;
        }
        
        // Load detailed properties for all pictures
        for (let i = 0; i < pictures.length; i++) {
          const picture = pictures[i];
          picture.load("name,id,left,top,width,height,zIndex,isSelected,altTextDescription,altTextTitle");
          
          // Try to load image data if available
          try {
            picture.load("imageData");
          } catch (e) {
            console.log("Couldn't load image data", e);
          }
        }
        await context.sync();
        
        // Check if any pictures are selected
        const selectedPictures = pictures.filter(pic => pic.isSelected);
        
        if (selectedPictures.length > 0) {
          // Found selected pictures!
          selectionInfo.pictures = selectedPictures;
          if (!selectionInfo.detectionMethod) {
            selectionInfo.detectionMethod = "picture-isSelected";
          }
        }
        
        // If no selected pictures found, check if any picture contains the last clicked position
        if (selectionInfo.pictures.length === 0 && hasMousePositionData) {
          const picturesAtPosition = [];
          
          for (let i = 0; i < pictures.length; i++) {
            const picture = pictures[i];
            
            // Check if the last clicked position is within this picture's bounds
            if (lastClickedPosition.x >= picture.left && 
                lastClickedPosition.x <= picture.left + picture.width &&
                lastClickedPosition.y >= picture.top && 
                lastClickedPosition.y <= picture.top + picture.height) {
              picturesAtPosition.push(picture);
            }
          }
          
          if (picturesAtPosition.length > 0) {
            // If multiple pictures overlap, take the one with highest z-index (top-most)
            const topPictures = picturesAtPosition.sort((a, b) => b.zIndex - a.zIndex);
            selectionInfo.pictures = topPictures;
            if (!selectionInfo.detectionMethod) {
              selectionInfo.detectionMethod = "picture-position";
            }
          }
        }
        
        // Store all pictures for reference
        if (selectionInfo.pictures.length === 0) {
          selectionInfo.allPictures = pictures;
        }
        
        resolve();
      } catch (error) {
        console.log("Picture-specific detection failed:", error);
        resolve();
      }
    });
  });
}

// Function to detect shapes directly
function detectShapesDirectly(selectionInfo) {
  return new Promise((resolve) => {
    PowerPoint.run(async (context) => {
      try {
        // Get the selected slides
        const selectedSlides = context.presentation.getSelectedSlides();
        selectedSlides.load("items");
        await context.sync();
        
        // Check if there are any selected slides
        if (!selectedSlides.items || selectedSlides.items.length === 0) {
          resolve();
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
              const shape = selectedShapes[i];
              shape.load("name,id,type,text,left,top,width,height,zIndex");
              
              // Load additional properties for pictures
              try {
                shape.load("altTextDescription,altTextTitle");
                
                // For pictures, load image-specific data
                if (shape.type === "Picture" || !shape.type) {
                  shape.load("imageData");
                }
              } catch (e) {
                console.log("Couldn't load extended properties for selected shape", e);
              }
            }
            await context.sync();
            
            // Add to selection info
            selectionInfo.shapes = selectedShapes;
            if (!selectionInfo.detectionMethod) {
              selectionInfo.detectionMethod = "direct-selection";
            }
            
            resolve();
            return;
          }
        } catch (selectionError) {
          console.log("Direct selection API not supported:", selectionError);
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
              
              // Add to selection info
              selectionInfo.shapes = selectedShapes;
              if (!selectionInfo.detectionMethod) {
                selectionInfo.detectionMethod = "active-view";
              }
              
              resolve();
              return;
            }
          }
        } catch (viewError) {
          console.log("Active view selection API not supported:", viewError);
        }
        
        // Approach 3: Check all shapes on the slide for isSelected property
        slide.load("shapes");
        await context.sync();
        
        if (!slide.shapes || !slide.shapes.items || slide.shapes.items.length === 0) {
          resolve();
          return;
        }
        
        // Load all shapes with their properties
        const shapes = slide.shapes.items;
        
        // Try to find selected shapes by checking isSelected property
        for (let i = 0; i < shapes.length; i++) {
          const shape = shapes[i];
          // Load all relevant properties including isSelected if available
          shape.load("name,id,type,text,left,top,width,height,zIndex,isSelected");
          
          // Load additional properties that might be available
          try {
            // These properties might not be available on all shapes or PowerPoint versions
            shape.load("altTextDescription,altTextTitle");
            
            // Try to load picture-specific properties if available
            if (shape.type === "Picture" || !shape.type) {
              shape.load("imageData");
            }
          } catch (e) {
            console.log("Couldn't load some extended properties", e);
          }
        }
        
        await context.sync();
        
        // Check if any shapes are marked as selected
        const selectedShapes = shapes.filter(shape => shape.isSelected);
        
        if (selectedShapes.length > 0) {
          // We found selected shapes!
          selectionInfo.shapes = selectedShapes;
          if (!selectionInfo.detectionMethod) {
            selectionInfo.detectionMethod = "isSelected";
          }
          
          resolve();
          return;
        }
        
        // Approach 4: Try to find shapes that were most recently modified
        try {
          // Get last modified shapes if available
          for (let i = 0; i < shapes.length; i++) {
            shapes[i].load("lastModified");
          }
          await context.sync();
          
          // Sort by last modified time if available
          const shapesWithModifiedTime = shapes.filter(s => s.lastModified);
          if (shapesWithModifiedTime.length > 0) {
            const sortedByModified = [...shapesWithModifiedTime].sort((a, b) => 
              new Date(b.lastModified) - new Date(a.lastModified)
            );
            
            if (sortedByModified.length > 0) {
              const mostRecentShape = sortedByModified[0];
              mostRecentShape.load("name,id,type,text,left,top,width,height,zIndex");
              await context.sync();
              
              selectionInfo.shapes = [mostRecentShape];
              if (!selectionInfo.detectionMethod) {
                selectionInfo.detectionMethod = "lastModified";
              }
              
              resolve();
              return;
            }
          }
        } catch (modifiedError) {
          console.log("Modified time detection failed:", modifiedError);
        }
        
        // Approach 5: Use spatial detection - find shape at last clicked position
        if (hasMousePositionData) {
          try {
            // Find shapes that contain the last clicked position
            const shapesAtPosition = [];
            
            for (let i = 0; i < shapes.length; i++) {
              const shape = shapes[i];
              
              // Check if the last clicked position is within this shape's bounds
              if (lastClickedPosition.x >= shape.left && 
                  lastClickedPosition.x <= shape.left + shape.width &&
                  lastClickedPosition.y >= shape.top && 
                  lastClickedPosition.y <= shape.top + shape.height) {
                shapesAtPosition.push(shape);
              }
            }
            
            if (shapesAtPosition.length > 0) {
              // If multiple shapes overlap, take the one with highest z-index (top-most)
              const topShapes = shapesAtPosition.sort((a, b) => b.zIndex - a.zIndex);
              selectionInfo.shapes = topShapes;
              if (!selectionInfo.detectionMethod) {
                selectionInfo.detectionMethod = "position";
              }
              
              resolve();
              return;
            }
          } catch (positionError) {
            console.log("Position-based detection failed:", positionError);
          }
        }
        
        // Approach 6: Try to detect the most recently clicked shape based on z-order
        try {
          // Sort shapes by z-index (higher z-index is typically the one on top that was clicked)
          const sortedShapes = [...shapes].sort((a, b) => b.zIndex - a.zIndex);
          
          if (sortedShapes.length > 0) {
            // Store all shapes for reference
            selectionInfo.allShapes = shapes;
            
            resolve();
            return;
          }
        } catch (zOrderError) {
          console.log("Z-order detection failed:", zOrderError);
        }
        
        // If we reach here, store all shapes for reference
        selectionInfo.allShapes = shapes;
        
        resolve();
      } catch (error) {
        console.error("Shape detection error:", error);
        resolve();
      }
    });
  });
}

// Function to display the selection information in a user-friendly format
function displaySelectionInfo(selectionInfo) {
  let message = "";
  
  // First, check if we have any selected elements
  const hasSelection = selectionInfo.text || 
                      (selectionInfo.shapes && selectionInfo.shapes.length > 0) || 
                      (selectionInfo.pictures && selectionInfo.pictures.length > 0);
  
  if (!hasSelection) {
    message = "No element is currently selected.\n\n";
    
    // Show all elements on the slide grouped by type
    const allShapes = selectionInfo.allShapes || [];
    const allPictures = selectionInfo.allPictures || [];
    
    if (allShapes.length > 0 || allPictures.length > 0) {
      message += `All elements on current slide (${allShapes.length + allPictures.length} total):\n\n`;
      
      // Group shapes by type
      const elementsByType = {};
      
      // Add shapes
      allShapes.forEach((shape) => {
        const type = shape.type || "Unknown";
        if (!elementsByType[type]) {
          elementsByType[type] = [];
        }
        elementsByType[type].push(shape);
      });
      
      // Add pictures (if not already included in shapes)
      if (allPictures.length > 0 && !elementsByType["Picture"]) {
        elementsByType["Picture"] = allPictures;
      }
      
      // Display elements grouped by type
      for (const type in elementsByType) {
        message += `== ${type} Elements (${elementsByType[type].length}) ==\n\n`;
        
        elementsByType[type].forEach((element, index) => {
          message += `${index + 1}. ${element.name || "Unnamed"} (ID: ${element.id})\n`;
          message += `   Position: (${element.left}, ${element.top})\n`;
          message += `   Size: ${element.width} x ${element.height}\n`;
          message += `   Z-Index: ${element.zIndex}\n\n`;
        });
      }
      
      message += "\nTips for selecting elements:\n";
      message += "1. Click directly on the element you want to examine\n";
      message += "2. Press the 'Get Info' button immediately after selecting\n";
      message += "3. For text elements, try selecting some text within the element\n";
      message += "4. Try using the 'Capture Current Position' button while hovering over an element";
    } else {
      message += "No elements found on the current slide.";
    }
    
    document.getElementById('item-subject').textContent = message;
    return;
  }
  
  // If we have a selection, show details about the selected element(s)
  message = "=== CURRENT SELECTION INFO ===\n\n";
  
  // If we have selected text, show it first
  if (selectionInfo.text) {
    message += "SELECTED TEXT:\n";
    message += `"${selectionInfo.text}"\n\n`;
  }
  
  // If we have selected pictures, show their details
  if (selectionInfo.pictures && selectionInfo.pictures.length > 0) {
    message += `SELECTED PICTURE${selectionInfo.pictures.length > 1 ? 'S' : ''}:\n`;
    
    selectionInfo.pictures.forEach((picture, index) => {
      message += getPictureDetailsString(picture, index);
    });
    
    message += "\n";
  }
  
  // If we have selected shapes, show their details
  if (selectionInfo.shapes && selectionInfo.shapes.length > 0) {
    // Filter out pictures that are already shown
    const nonPictureShapes = selectionInfo.shapes.filter(shape => 
      shape.type !== "Picture" || 
      !selectionInfo.pictures || 
      !selectionInfo.pictures.some(pic => pic.id === shape.id)
    );
    
    if (nonPictureShapes.length > 0) {
      message += `SELECTED SHAPE${nonPictureShapes.length > 1 ? 'S' : ''}:\n`;
      
      nonPictureShapes.forEach((shape, index) => {
        message += getDetailedShapeInfo(shape, index);
      });
    }
  }
  
  // Add detection method information
  if (selectionInfo.detectionMethod) {
    message += "\nDetection method: " + selectionInfo.detectionMethod;
  }
  
  document.getElementById('item-subject').textContent = message;
}

// Helper function to get detailed information about a picture
function getPictureDetailsString(picture, index) {
  let details = `${index + 1}. ${picture.name || "Unnamed Picture"} (ID: ${picture.id})\n`;
  details += `   Type: Picture/Image\n`;
  details += `   Position: (${picture.left}, ${picture.top})\n`;
  details += `   Size: ${picture.width} x ${picture.height}\n`;
  details += `   Z-Index: ${picture.zIndex}\n`;
  
  // Add picture-specific details
  if (picture.altTextDescription) {
    details += `   Alt Text: ${picture.altTextDescription}\n`;
  }
  
  if (picture.altTextTitle) {
    details += `   Alt Text Title: ${picture.altTextTitle}\n`;
  }
  
  if (picture.imageData && picture.imageData.format) {
    details += `   Image Format: ${picture.imageData.format}\n`;
  }
  
  details += "\n";
  return details;
}

// Helper function to get detailed information about a shape
function getDetailedShapeInfo(shape, index) {
  let details = `${index + 1}. ${shape.name || "Unnamed"} (ID: ${shape.id})\n`;
  details += `   Type: ${shape.type}\n`;
  
  // Add specific details based on shape type
  switch (shape.type) {
    case "Picture":
      details += `   Element Type: Image/Picture\n`;
      
      // Load additional picture-specific properties if available
      try {
        if (shape.imageData) {
          details += `   Image Format: ${shape.imageData.format || "Unknown"}\n`;
        }
        
        if (shape.altTextDescription) {
          details += `   Alt Text: ${shape.altTextDescription}\n`;
        }
        
        if (shape.altTextTitle) {
          details += `   Alt Text Title: ${shape.altTextTitle}\n`;
        }
      } catch (e) {
        console.log("Couldn't load some picture properties", e);
      }
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
