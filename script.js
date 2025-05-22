document.getElementById('uploadForm').addEventListener('submit', async function(e) {
    e.preventDefault();
    
    // Получаем значения из формы
    const formData = new FormData(this);
    const data = {};
    formData.forEach((value, key) => {
        data[key] = value;
    });

    try {
        // Создаем новый документ
        const response = await fetch('https://raw.githubusercontent.com/AgniStore/titulnik/main/tit.docx');
        if (!response.ok) {
            throw new Error('Не удалось загрузить шаблон документа');
        }
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

        // Показываем сообщение об успехе
        alert('Документ успешно скачан!');
    } catch (error) {
        console.error('Ошибка:', error);
        alert('Произошла ошибка при скачивании документа. Пожалуйста, попробуйте позже.');
    }
}); 