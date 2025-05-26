Office.onReady((info) => {
    if (info.host === Office.HostType.PowerPoint) {
        document.getElementById('run-button').onclick = run;
    }
});

async function run() {
    try {
        await PowerPoint.run(async (context) => {
            // Get the current slide
            const slide = context.presentation.getSelectedSlides().getFirst();
            slide.load("id");
            await context.sync();

            // Get selected shapes on the slide
            const selectedShapes = slide.shapes.getSelected();
            // Load more properties for detailed information
            selectedShapes.load("items/name,items/id,items/type,items/left,items/top,items/width,items/height,items/textFrame/text");
            await context.sync();

            let message = "";
            if (selectedShapes.items.length > 0) {
                message = "Selected element(s):\n";
                
                selectedShapes.items.forEach(shape => {
                    message += `- Name: ${shape.name}\n`;
                    message += `  Type: ${shape.type}\n`;
                    message += `  Position: (${shape.left}, ${shape.top})\n`;
                    message += `  Size: ${shape.width} x ${shape.height}\n`;
                    
                    // Check if shape has text
                    if (shape.textFrame && shape.textFrame.text) {
                        message += `  Text: "${shape.textFrame.text}"\n`;
                    }
                    
                    message += "\n";
                });
            } else {
                message = "No elements are selected. Please select a shape on the slide.";
            }
            document.getElementById('message').textContent = message;
        });
    } catch (error) {
        console.error(error);
        document.getElementById('message').textContent = 'Error: ' + error;
    }
} 