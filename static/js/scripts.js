document.getElementById('uploadForm').addEventListener('submit', async function (e) {
    e.preventDefault();
    const formData = new FormData(this);

    const responseDiv = document.getElementById('response');
    responseDiv.innerHTML = 'Uploading and processing...';

    try {
        const response = await fetch('/upload', {
            method: 'POST',
            body: formData,
        });

        const result = await response.json();

        if (response.ok) {
            responseDiv.innerHTML = `<p style="color: green;">${result.message}</p>`;
        } else {
            responseDiv.innerHTML = `<p style="color: red;">Error: ${result.error}</p>`;
        }
    } catch (err) {
        responseDiv.innerHTML = `<p style="color: red;">Unexpected error: ${err.message}</p>`;
    }
});
