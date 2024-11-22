document.addEventListener('DOMContentLoaded', function() {
    // Get form elements
    const form = document.getElementById('uploadForm');
    const fileInput = document.getElementById('documentFile');
    
    // Only add event listeners if elements exist
    if (form) {
        const processingStatus = document.getElementById('processingStatus');
        const downloadSection = document.getElementById('downloadSection');

        form.addEventListener('submit', async function(e) {
            e.preventDefault();

            // Form validation
            if (!form.checkValidity()) {
                e.stopPropagation();
                form.classList.add('was-validated');
                return;
            }

            // Prepare form data
            const formData = new FormData(form);
            
            // Show processing status
            processingStatus.classList.remove('d-none');
            downloadSection.classList.add('d-none');

            try {
                const response = await fetch('/process', {
                    method: 'POST',
                    body: formData
                });

                // Check if response is redirect (HTTP 302)
                if (response.redirected) {
                    window.location.href = response.url;
                    return;
                }

                const data = await response.json();

                if (response.ok) {
                    // Show download section
                    downloadSection.classList.remove('d-none');
                } else {
                    alert(data.error || 'An error occurred while processing the document.');
                }
            } catch (error) {
                console.error('Processing error:', error);
                alert('An error occurred while processing the document. Please try again.');
            } finally {
                processingStatus.classList.add('d-none');
            }
        });
    }

    // Add file input change handler for better UX
    if (fileInput) {
        fileInput.addEventListener('change', function() {
            const fileName = this.files[0]?.name;
            const label = document.querySelector(`label[for="${this.id}"]`);
            if (label) {
                label.textContent = fileName || 'Select Word Document';
            }
        });
    }
}); 