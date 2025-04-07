document.addEventListener('DOMContentLoaded', function() {
    const dropZone = document.getElementById('dropZone');
    const fileInput = document.getElementById('fileInput');
    const fileList = document.getElementById('fileList');
    const processButton = document.getElementById('processButton');
    const progressArea = document.getElementById('progressArea');
    const progressBar = document.getElementById('progressBar');
    const resultArea = document.getElementById('resultArea');
    const resultMessage = document.getElementById('resultMessage');
    const downloadButton = document.getElementById('downloadButton');

    let files = [];

    // 拖放功能
    dropZone.addEventListener('dragover', (e) => {
        e.preventDefault();
        dropZone.style.borderColor = '#4CAF50';
    });

    dropZone.addEventListener('dragleave', (e) => {
        e.preventDefault();
        dropZone.style.borderColor = '#ccc';
    });

    dropZone.addEventListener('drop', (e) => {
        e.preventDefault();
        dropZone.style.borderColor = '#ccc';
        
        const newFiles = Array.from(e.dataTransfer.files).filter(
            file => file.name.endsWith('.xlsx') || file.name.endsWith('.csv')
        );
        
        if (newFiles.length > 0) {
            files = newFiles;
            updateFileList();
        }
    });

    // 点击上传
    dropZone.addEventListener('click', () => {
        fileInput.click();
    });

    fileInput.addEventListener('change', (e) => {
        files = Array.from(e.target.files);
        updateFileList();
    });

    function updateFileList() {
        fileList.innerHTML = '';
        files.forEach(file => {
            const fileItem = document.createElement('div');
            fileItem.className = 'file-item';
            fileItem.innerHTML = `
                <span>${file.name}</span>
                <button onclick="removeFile('${file.name}')" class="remove-button">删除</button>
            `;
            fileList.appendChild(fileItem);
        });
        processButton.disabled = files.length === 0;
    }

    window.removeFile = function(fileName) {
        files = files.filter(file => file.name !== fileName);
        updateFileList();
    };

    // 处理文件
    processButton.addEventListener('click', async () => {
        if (files.length === 0) return;

        const formData = new FormData();
        files.forEach(file => {
            formData.append('files[]', file);
        });

        processButton.disabled = true;
        progressArea.style.display = 'block';
        progressBar.style.width = '50%';

        try {
            const response = await fetch('/upload', {
                method: 'POST',
                body: formData
            });

            const result = await response.json();

            if (response.ok) {
                progressBar.style.width = '100%';
                resultArea.style.display = 'block';
                resultMessage.className = 'success';
                resultMessage.textContent = '文件处理成功！';
                downloadButton.style.display = 'block';
            } else {
                throw new Error(result.error);
            }
        } catch (error) {
            resultArea.style.display = 'block';
            resultMessage.className = 'error';
            resultMessage.textContent = `错误：${error.message}`;
            downloadButton.style.display = 'none';
        } finally {
            progressArea.style.display = 'none';
            processButton.disabled = false;
        }
    });

    // 下载处理后的文件
    downloadButton.addEventListener('click', () => {
        window.location.href = '/download';
    });

    // 添加文件比较相关的代码
    const compareButton = document.getElementById('compareButton');
    const file1Input = document.getElementById('file1');
    const file2Input = document.getElementById('file2');
    const compareProgress = document.getElementById('compareProgress');
    const compareResult = document.getElementById('compareResult');
    const compareMessage = document.getElementById('compareMessage');
    const downloadComparisonButton = document.getElementById('downloadComparisonButton');
    
    compareButton.addEventListener('click', async () => {
        if (!file1Input.files[0] || !file2Input.files[0]) {
            alert('请选择两个文件进行比较');
            return;
        }

        const formData = new FormData();
        formData.append('file1', file1Input.files[0]);
        formData.append('file2', file2Input.files[0]);

        compareButton.disabled = true;
        compareProgress.style.display = 'block';
        compareResult.style.display = 'none';

        try {
            const response = await fetch('/compare', {
                method: 'POST',
                body: formData
            });

            const result = await response.json();

            if (response.ok) {
                compareResult.style.display = 'block';
                compareMessage.className = 'success';
                compareMessage.textContent = '文件比较成功！';
                downloadComparisonButton.style.display = 'block';
                downloadComparisonButton.onclick = () => {
                    window.location.href = `/download_comparison/${result.timestamp}`;
                };
            } else {
                throw new Error(result.error);
            }
        } catch (error) {
            compareResult.style.display = 'block';
            compareMessage.className = 'error';
            compareMessage.textContent = `错误：${error.message}`;
            downloadComparisonButton.style.display = 'none';
        } finally {
            compareProgress.style.display = 'none';
            compareButton.disabled = false;
        }
    });
});
