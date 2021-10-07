# prescription_flask

Criador de receitas especiais em Flask para compilação com "pywebview".

Recomendo utilizar Python 3.7 para evitar conflitos com o pythonnet.

O erro da biblioteca "docx2pdf" com o hook do pyinstaller pode ser corrigido criando-se um arquivo com o nome "hook-docx2pdf.py", no diretório c:\users\'usuario'\appdata\local\programs\python\python37\Lib\site-packages\PyInstaller\hooks", com o seguinte conteúdo:

from PyInstaller.utils.hooks import collect_all
datas, binaries, hiddenimports = collect_all('docx2pdf')
