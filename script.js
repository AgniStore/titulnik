document.getElementById('uploadForm').addEventListener('submit', async function(e) {
    e.preventDefault();
    
    // Получаем значения из формы
    const formData = new FormData(this);
    const data = {};
    formData.forEach((value, key) => {
        data[key] = value;
    });

    // Создаем новый документ
    const response = await fetch('https://raw.githubusercontent.com/AgniStore/titulnik/main/tit.docx', {
        method: 'GET'
    });
    const blob = await response.blob();

    // Создаем временную ссылку для скачивания
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'filled_document.docx';
    document.body.appendChild(a);
    a.click();
    window.URL.revokeObjectURL(url);
    document.body.removeChild(a);
}); 