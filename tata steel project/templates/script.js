function downloadExcel() {
    var data = [];
    $('#dataDisplay tbody tr').each(function() {
        var row = $(this);
        data.push({
            name: row.find('td:eq(0)').text(),
            phone: row.find('td:eq(1)').text(),
            email: row.find('td:eq(2)').text()
        });
    });

    $.ajax({
        type: 'POST',
        url: '/download_excel',
        contentType: 'application/json',
        data: JSON.stringify({data: data}),
        xhrFields: {
            responseType: 'blob'
        },
        success: function(response, status, xhr) {
            var a = document.createElement('a');
            var url = window.URL.createObjectURL(response);
            a.href = url;
            a.download = 'download.xlsx';  // Make sure the filename is provided here if not in the header
            document.body.appendChild(a);
            a.click();
            a.remove();
            window.URL.revokeObjectURL(url);
        },
        error: function(xhr, status, error) {
            alert('Failed to download Excel: ' + error);
        }
    });
}
