document.getElementById('lessonForm').addEventListener('submit', async function(e) {
    e.preventDefault();

    // Get input elements and button state
    const subject = document.getElementById('subject').value;
    const classLevel = document.getElementById('classLevel').value;
    const topic = document.getElementById('topic').value.trim();
    const submitBtn = document.getElementById('submitBtn');

    if (!topic) {
        alert("Please enter a valid topic.");
        return;
    }

    // Adjust UI state during generation latency
    submitBtn.disabled = true;
    submitBtn.textContent = 'Generating & Downloading...';

    try {
        // Post structured JSON data to the FastAPI endpoint
        const response = await fetch('http://127.0.0.1:8001/api/v1/export-lesson-note', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
            },
            body: JSON.stringify({
                subject: subject,
                class_level: classLevel,
                topic: topic
            }),
        });

        if (!response.ok) {
            throw new Error(`Server returned status: ${response.status}`);
        }

        // Catch the streaming document response as a raw Binary Blob
        const blob = await response.blob();
        
        // Generate a localized DOM reference address pointing to the file data
        const downloadUrl = window.URL.createObjectURL(blob);
        
        // Create an ephemeral layout element to force download execution behavior
        const hiddenAnchor = document.createElement('a');
        hiddenAnchor.href = downloadUrl;
        
        // Sanitize spaces inside filenames to prevent formatting drops
        const safeTopic = topic.replace(/\s+/g, '_');
        hiddenAnchor.setAttribute('download', `LessonNote_${subject}_${classLevel}_${safeTopic}.docx`);
        
        // Execute document extraction
        document.body.appendChild(hiddenAnchor);
        hiddenAnchor.click();
        
        // Reclaim local context execution memory parameters
        hiddenAnchor.remove();
        window.URL.revokeObjectURL(downloadUrl);

    } catch (error) {
        console.error('Generation pipeline tracking error:', error);
        alert('Could not generate document. Check that your Python backend main.py server is running.');
    } finally {
        // Re-enable interactive form elements
        submitBtn.disabled = false;
        submitBtn.textContent = 'Generate Note';
    }
});