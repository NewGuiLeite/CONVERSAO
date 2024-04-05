document.getElementById('convertButton').addEventListener('click', function() {
    var input = document.getElementById('excelFile');
    var file = input.files[0];
    if (!file) {
      alert('Por favor, selecione um arquivo Excel.');
      return;
    }
  
    var reader = new FileReader();
    reader.onload = function(e) {
      var data = new Uint8Array(e.target.result);
      var workbook = XLSX.read(data, { type: 'array' });
  
      var updateSQL = '';
      workbook.SheetNames.forEach(function(sheetName) {
        var sheet = workbook.Sheets[sheetName];
        var range = XLSX.utils.decode_range(sheet['!ref']);
  
        for (var rowNum = range.s.r; rowNum <= range.e.r; rowNum++) {
          var codigo = sheet['A' + (rowNum + 1)] ? sheet['A' + (rowNum + 1)].v : '';
          var saldo = sheet['B' + (rowNum + 1)] ? sheet['B' + (rowNum + 1)].v : '';
  
          if (codigo && saldo) {
            updateSQL += "UPDATE produto SET pro_saldo = " + saldo + " WHERE pro_cod = '" + codigo + "';\n";
          }
        }
      });

      // Adiciona o COMMIT WORK no final das instruções SQL
      updateSQL += "COMMIT WORK;\n";

    
      document.getElementById('output').innerText = updateSQL;
    };
    reader.readAsArrayBuffer(file);
  });
  

  
document.getElementById('copyButton').addEventListener('click', function() {
  var output = document.getElementById('output');
  var range = document.createRange();
  range.selectNode(output);
  window.getSelection().removeAllRanges();
  window.getSelection().addRange(range);
  document.execCommand('copy');
  window.getSelection().removeAllRanges();
  alert('Código copiado para a área de transferência!');
});

document.getElementById('saveButton').addEventListener('click', function() {
  var output = document.getElementById('output').innerText;
  var blob = new Blob([output], { type: 'text/plain' });
  var url = window.URL.createObjectURL(blob);
  var a = document.createElement('a');
  a.href = url;
  a.download = 'codigo_sql.txt';
  a.click();
  window.URL.revokeObjectURL(url);
});
