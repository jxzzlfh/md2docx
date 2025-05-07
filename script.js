document.addEventListener('DOMContentLoaded', function() {
    // 初始化Markdown解析器
    const md = window.markdownit({
        html: true,
        breaks: true,
        linkify: true,
        typographer: true
    });

    // 初始化HTML到Markdown转换器
    const turndownService = new TurndownService({
        headingStyle: 'atx',
        codeBlockStyle: 'fenced'
    });

    // 获取DOM元素
    const importBtn = document.getElementById('import-btn');
    const editBtn = document.getElementById('edit-btn');
    const previewBtn = document.getElementById('preview-btn');
    const convertBtn = document.getElementById('convert-btn');
    
    const importSection = document.getElementById('import-section');
    const editSection = document.getElementById('edit-section');
    const previewSection = document.getElementById('preview-section');
    const convertSection = document.getElementById('convert-section');
    
    const fileInput = document.getElementById('file-input');
    const fileUploadBox = document.querySelector('.file-upload-box');
    const clipboardImportBtn = document.getElementById('clipboard-import');
    const manualPasteBtn = document.getElementById('manual-paste');
    const editor = document.getElementById('editor');
    const previewContent = document.getElementById('preview-content');
    
    const lineCount = document.getElementById('line-count');
    const wordCount = document.getElementById('word-count');
    const charCount = document.getElementById('char-count');
    
    const downloadMdBtn = document.getElementById('download-md');
    const downloadDocxBtn = document.getElementById('download-docx');
    
    // 工具栏按钮
    const toolbarButtons = document.querySelectorAll('.editor-toolbar button');

    // 切换活动部分
    function setActiveSection(section) {
        // 移除所有按钮的活动状态
        [importBtn, editBtn, previewBtn, convertBtn].forEach(btn => {
            btn.classList.remove('active');
        });
        
        // 隐藏所有部分
        [importSection, editSection, previewSection, convertSection].forEach(sec => {
            sec.classList.remove('active');
        });
        
        // 设置活动部分和按钮
        if (section === 'import') {
            importBtn.classList.add('active');
            importSection.classList.add('active');
        } else if (section === 'edit') {
            editBtn.classList.add('active');
            editSection.classList.add('active');
            updateEditorStats();
        } else if (section === 'preview') {
            previewBtn.classList.add('active');
            previewSection.classList.add('active');
            renderPreview();
        } else if (section === 'convert') {
            convertBtn.classList.add('active');
            convertSection.classList.add('active');
        }
    }

    // 按钮点击事件
    importBtn.addEventListener('click', () => setActiveSection('import'));
    editBtn.addEventListener('click', () => setActiveSection('edit'));
    previewBtn.addEventListener('click', () => setActiveSection('preview'));
    convertBtn.addEventListener('click', () => setActiveSection('convert'));

    // 记录剪贴板API状态的变量
    let clipboardAPIFailed = false;

    // 从剪贴板导入功能
    clipboardImportBtn.addEventListener('click', async () => {
        // 如果之前尝试失败过，直接使用手动粘贴方式
        if (clipboardAPIFailed) {
            manualPasteBtn.click();
            return;
        }
        
        // 先检查剪贴板API是否可用
        if (!navigator.clipboard || !navigator.clipboard.readText) {
            clipboardAPIFailed = true;
            manualPasteBtn.click();
            return;
        }
        
        try {
            // 添加加载提示
            showLoadingMessage('正在读取剪贴板...');
            
            // 使用超时处理，避免长时间等待
            const timeoutPromise = new Promise((_, reject) => 
                setTimeout(() => reject(new Error('剪贴板读取超时')), 3000)
            );
            
            // 尝试读取剪贴板内容，并添加超时处理
            const text = await Promise.race([
                navigator.clipboard.readText().catch(e => {
                    console.warn('读取剪贴板时出现警告:', e);
                    return '';
                }),
                timeoutPromise
            ]).catch(error => {
                console.warn('剪贴板读取失败或超时:', error);
                clipboardAPIFailed = true;
                return '';
            });
            
            // 隐藏加载提示
            hideLoadingMessage();
            
            // 如果没有获取到内容，或API调用失败
            if (!text || text.trim() === '') {
                if (clipboardAPIFailed) {
                    // 如果API调用失败，直接使用备选方案
                    manualPasteBtn.click();
                } else {
                    // 内容为空，显示提示
                    alert('剪贴板内容为空');
                }
                return;
            }
            
            // 设置内容并切换到编辑模式
            editor.value = text;
            setActiveSection('edit');
            updateEditorStats();
        } catch (error) {
            // 标记API调用失败
            clipboardAPIFailed = true;
            
            // 隐藏加载提示
            hideLoadingMessage();
            
            console.error('读取剪贴板失败:', error);
            
            // 使用备选方案
            manualPasteBtn.click();
        }
    });

    // 手动粘贴按钮功能
    manualPasteBtn.addEventListener('click', () => {
        // 直接切换到编辑模式
        setActiveSection('edit');
        
        // 清空编辑器内容
        editor.value = '';
        
        // 聚焦到编辑器并显示提示
        editor.focus();
        alert('请使用Ctrl+V（或Command+V）直接粘贴内容到编辑区');
    });

    // 拖放和文件选择功能
    fileUploadBox.addEventListener('click', () => {
        fileInput.click();
    });

    fileUploadBox.addEventListener('dragover', (e) => {
        e.preventDefault();
        fileUploadBox.classList.add('highlight');
    });

    fileUploadBox.addEventListener('dragleave', () => {
        fileUploadBox.classList.remove('highlight');
    });

    fileUploadBox.addEventListener('drop', (e) => {
        e.preventDefault();
        fileUploadBox.classList.remove('highlight');
        
        if (e.dataTransfer.files.length) {
            handleFile(e.dataTransfer.files[0]);
        }
    });

    fileInput.addEventListener('change', (e) => {
        if (e.target.files.length) {
            handleFile(e.target.files[0]);
        }
    });

    // 处理导入的文件
    function handleFile(file) {
        if (file.type !== 'text/markdown' && !file.name.endsWith('.md') && !file.name.endsWith('.markdown')) {
            alert('请选择Markdown文件（.md或.markdown）');
            return;
        }
        
        const reader = new FileReader();
        reader.onload = (e) => {
            editor.value = e.target.result;
            setActiveSection('edit');
            updateEditorStats();
        };
        reader.readAsText(file);
    }

    // 编辑器工具栏功能
    toolbarButtons.forEach(button => {
        button.addEventListener('click', () => {
            const action = button.getAttribute('data-action');
            handleToolbarAction(action);
        });
    });

    function handleToolbarAction(action) {
        const selectionStart = editor.selectionStart;
        const selectionEnd = editor.selectionEnd;
        const selectedText = editor.value.substring(selectionStart, selectionEnd);
        
        let replacement = '';
        let cursorOffset = 0;

        switch (action) {
            case 'bold':
                replacement = `**${selectedText}**`;
                cursorOffset = selectedText ? 0 : -2;
                break;
            case 'italic':
                replacement = `*${selectedText}*`;
                cursorOffset = selectedText ? 0 : -1;
                break;
            case 'heading':
                replacement = `# ${selectedText}`;
                cursorOffset = selectedText ? 0 : 0;
                break;
            case 'quote':
                replacement = `> ${selectedText}`;
                cursorOffset = selectedText ? 0 : 0;
                break;
            case 'ul':
                replacement = `* ${selectedText}`;
                cursorOffset = selectedText ? 0 : 0;
                break;
            case 'ol':
                replacement = `1. ${selectedText}`;
                cursorOffset = selectedText ? 0 : 0;
                break;
            case 'link':
                replacement = `[${selectedText}](url)`;
                cursorOffset = selectedText ? -1 : -5;
                break;
            case 'image':
                replacement = `![${selectedText}](image-url)`;
                cursorOffset = selectedText ? -1 : -11;
                break;
            case 'code':
                replacement = `\`\`\`\n${selectedText}\n\`\`\``;
                cursorOffset = selectedText ? 0 : -4;
                break;
            case 'fullscreen':
                editor.parentElement.classList.toggle('fullscreen');
                return;
            case 'help':
                alert('Markdown编辑器帮助：\n\n' +
                      '加粗 - 使文字加粗\n' +
                      '斜体 - 使文字倾斜\n' +
                      '标题 - 添加标题\n' +
                      '引用 - 添加引用文字\n' +
                      '无序列表 - 创建无序列表\n' +
                      '有序列表 - 创建有序列表\n' +
                      '链接 - 添加超链接\n' +
                      '图片 - 插入图片\n' +
                      '代码块 - 添加代码块\n' +
                      '全屏编辑 - 切换全屏编辑模式\n' +
                      '帮助 - 显示此帮助信息');
                return;
        }

        // 保存当前滚动位置
        const scrollTop = editor.scrollTop;
        
        // 修改文本内容
        editor.value = editor.value.substring(0, selectionStart) + replacement + editor.value.substring(selectionEnd);
        
        // 设置新的光标位置，基于原始位置加上替换文本的长度差异
        const newCursorPos = selectionStart + replacement.length - (selectionEnd - selectionStart) + cursorOffset;
        
        // 明确地将焦点集中到编辑器，确保光标可见
        editor.focus();
        
        // 设置新的光标位置
        editor.selectionStart = newCursorPos;
        editor.selectionEnd = newCursorPos;
        
        // 恢复原始滚动位置
        editor.scrollTop = scrollTop;
        
        updateEditorStats();
    }

    // 更新编辑器统计信息
    function updateEditorStats() {
        const text = editor.value;
        
        // 计算汉字数量
        const chineseCharCount = (text.match(/[\u4e00-\u9fa5]/g) || []).length;
        
        // 更新统计信息，只显示汉字计数
        charCount.textContent = `字数: ${chineseCharCount}`;
        
        // 隐藏其他统计信息
        lineCount.style.display = 'none';
        wordCount.style.display = 'none';
    }

    // 监听编辑器内容变化
    editor.addEventListener('input', updateEditorStats);
    editor.addEventListener('keyup', updateEditorStats);

    // 渲染预览
    function renderPreview() {
        const content = editor.value;
        const htmlContent = md.render(content);
        previewContent.innerHTML = htmlContent;
    }

    // 下载功能
    downloadMdBtn.addEventListener('click', () => {
        const mdContent = editor.value;
        if (mdContent.trim()) {
            downloadFile(mdContent, 'document.md', 'text/markdown');
        } else {
            alert('没有内容可下载！');
        }
    });
    
    downloadDocxBtn.addEventListener('click', () => {
        const mdContent = editor.value;
        if (mdContent.trim()) {
            // 显示加载消息
            showLoadingMessage('正在转换为DOCX...');
            
            // 渲染HTML
            const htmlContent = md.render(mdContent);
            
            // 创建完整的HTML文档，并添加基础样式
            const styledHtml = createStyledHtml(htmlContent);
            
            // 转换为DOCX
            if (window.convertHtmlToDocx) {
                window.convertHtmlToDocx(styledHtml, (blob) => {
                    hideLoadingMessage();
                    if (blob) {
                        console.log("DOCX转换成功，保存文件");
                        saveAs(blob, 'document.docx');
                    } else {
                        console.error("DOCX转换返回null");
                        // 转换失败处理策略
                        if (confirm('DOCX转换失败。您可以选择：\n1. 下载为HTML格式并在Word中打开\n2. 尝试直接在Word中粘贴预览内容\n\n是否下载HTML格式？')) {
                            // 下载HTML格式
                            downloadHtmlFile(styledHtml);
                        } else {
                            // 切换到预览，让用户复制内容
                            alert('已切换到预览模式，您可以全选内容后复制，然后粘贴到Word中。');
                            setActiveSection('preview');
                        }
                    }
                });
            } else {
                console.error("convertHtmlToDocx函数不存在");
                hideLoadingMessage();
                // 当转换函数不可用时
                const fallbackOptions = '无法使用DOCX转换功能。您可以：\n' +
                                      '1. 下载为HTML格式并在Word中打开\n' +
                                      '2. 查看预览并手动复制内容到Word\n\n' +
                                      '是否下载HTML格式？';
                                      
                if (confirm(fallbackOptions)) {
                    downloadHtmlFile(styledHtml);
                } else {
                    setActiveSection('preview');
                }
            }
        } else {
            alert('没有内容可下载！');
        }
    });
    
    // 创建带样式的HTML（优化Word转换效果）
    function createStyledHtml(htmlContent) {
        return `<!DOCTYPE html>
        <html>
        <head>
            <meta charset="UTF-8">
            <style>
                body {
                    font-family: 'SimSun', 'Times New Roman', serif;
                    font-size: 12pt;
                    line-height: 1.3;
                    margin: 2cm;
                }
                h1, h2, h3, h4, h5, h6 {
                    margin-top: 12pt;
                    margin-bottom: 6pt;
                    font-weight: bold;
                }
                h1 { font-size: 18pt; }
                h2 { font-size: 16pt; }
                h3 { font-size: 14pt; }
                h4, h5, h6 { font-size: 12pt; }
                p {
                    margin-top: 0;
                    margin-bottom: 6pt;
                }
                ul, ol {
                    margin-top: 3pt;
                    margin-bottom: 3pt;
                    padding-left: 20pt;
                }
                li {
                    margin-top: 2pt;
                    margin-bottom: 2pt;
                }
                ul li {
                    list-style-type: disc;
                }
                ol li {
                    list-style-type: decimal;
                }
                ul ul li {
                    list-style-type: circle;
                }
                ol ol li {
                    list-style-type: lower-alpha;
                }
                ul ul ul li {
                    list-style-type: square;
                }
                li > ul, li > ol {
                    margin-top: 2pt;
                    margin-bottom: 2pt;
                }
                table {
                    border-collapse: collapse;
                    width: 100%;
                    margin-top: 8pt;
                    margin-bottom: 8pt;
                }
                th, td {
                    border: 1px solid #ddd;
                    padding: 4pt;
                    text-align: left;
                }
                th {
                    background-color: #f2f2f2;
                    font-weight: bold;
                }
                blockquote {
                    margin: 8pt 0;
                    padding-left: 12pt;
                    border-left: 4px solid #ddd;
                    color: #666;
                }
                pre {
                    background-color: #f5f5f5;
                    padding: 8pt;
                    overflow-x: auto;
                    font-family: 'Courier New', monospace;
                    margin: 8pt 0;
                }
                code {
                    font-family: 'Courier New', monospace;
                    background-color: #f5f5f5;
                    padding: 2pt 4pt;
                    border-radius: 3px;
                }
                img {
                    max-width: 100%;
                    height: auto;
                }
            </style>
        </head>
        <body>
            ${htmlContent}
        </body>
        </html>`;
    }
    
    // HTML文件下载方法
    function downloadHtmlFile(htmlContent) {
        const blob = new Blob([htmlContent], { type: 'text/html' });
        
        // 使用.doc扩展名而不是.html，以便系统默认用Word打开
        saveAs(blob, 'document.doc');
        
        setTimeout(() => {
            alert('已将文档导出为兼容Word的HTML格式。\n\n此文件使用.doc扩展名，大多数系统将默认用Word打开。\n\n打开后，您可以保存为.docx格式并进行进一步编辑。');
        }, 500);
    }

    // 显示加载消息
    function showLoadingMessage(message) {
        const loadingDiv = document.createElement('div');
        loadingDiv.id = 'loading-message';
        loadingDiv.style.position = 'fixed';
        loadingDiv.style.top = '50%';
        loadingDiv.style.left = '50%';
        loadingDiv.style.transform = 'translate(-50%, -50%)';
        loadingDiv.style.background = 'rgba(0, 0, 0, 0.7)';
        loadingDiv.style.color = 'white';
        loadingDiv.style.padding = '20px 30px';
        loadingDiv.style.borderRadius = '5px';
        loadingDiv.style.zIndex = '1000';
        loadingDiv.textContent = message;
        document.body.appendChild(loadingDiv);
    }

    // 隐藏加载消息
    function hideLoadingMessage() {
        const loadingDiv = document.getElementById('loading-message');
        if (loadingDiv) {
            document.body.removeChild(loadingDiv);
        }
    }

    // 下载文件
    function downloadFile(content, filename, contentType) {
        const blob = new Blob([content], { type: contentType });
        const url = URL.createObjectURL(blob);
        
        const a = document.createElement('a');
        a.href = url;
        a.download = filename;
        a.click();
        
        URL.revokeObjectURL(url);
    }

    // 初始化
    setActiveSection('import');
}); 