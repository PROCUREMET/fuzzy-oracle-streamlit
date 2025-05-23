from kivy.app import App
from kivy.lang import Builder

class meuAplicativo(App):
    def build(self):
        # Descarregar o arquivo .kv e carregá-lo novamente
        Builder.unload_file("tela.kv")  # Descarrega o arquivo .kv
        return Builder.load_file("tela.kv")  # Carrega novamente o arquivo .kv
    
# Executa o aplicativo
meuAplicativo().run()