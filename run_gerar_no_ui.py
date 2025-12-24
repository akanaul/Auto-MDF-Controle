# Wrapper para executar gerar_planilha.py sem UI interativa (mocka tkinter dialogs)
import runpy

# Mock simples para evitar caixas de diálogo
class Dummy:
    @staticmethod
    def askstring(title, prompt, **kwargs):
        return 'NAO INFORMADO'

    @staticmethod
    def showinfo(*args, **kwargs):
        return None

    @staticmethod
    def showerror(*args, **kwargs):
        return None

# Injetar mocks antes de importar o script
import sys
import types
mock_simpledialog = types.SimpleNamespace(askstring=Dummy.askstring)
mock_messagebox = types.SimpleNamespace(showinfo=Dummy.showinfo, showerror=Dummy.showerror)

# Por segurança, colocar no sys.modules para que quando o script importar tkinter.simpledialog
# e tkinter.messagebox ele receba nossas versões simples
import tkinter
sys.modules['tkinter.simpledialog'] = mock_simpledialog
sys.modules['tkinter.messagebox'] = mock_messagebox

# Executar o script principal
runpy.run_path('gerar_planilha.py', run_name='__main__')
