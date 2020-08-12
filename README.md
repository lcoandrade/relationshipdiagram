# relationshipdiagram
Relationship Diagram (Directed Graph)

Video com demonstração do uso: https://youtu.be/6xwjbfyR9Zg

Para executar esse script, é necessário:

 * Python 3 instalado e no PATH
   - Disponível em https://www.python.org/downloads/
   - No windows, marcar a inclusão no PATH durante a instalação
   - Instruções detalhadas para instalação do Python no Windows: https://python.org.br/instalacao-windows/
   - Instruções detalhadas para instalação do Python no macOS:https://www.python.org/downloads/mac-osx/
 
 * Graphviz instalado e no PATH
   - No Windows:
      - Disponível em https://graphviz.org/download/
      - No windows, marcar a inclusão no PATH durante a instalação. Caso falhe, consultar as instruções detalhadas abaixo.
      - No windows, após instalar, executar o comando "dot -c" como administrador (necessário apenas uma vez)
      - Instruções detalhadas para instalação do Graphviz no Windows: https://www.slideshare.net/cristianorolim1/tutorial-de-instalao-do-graphviz-no-windows-10
   - No macOS:
      - Executar no terminal: xcode-select --install
      - Instalar o homebrew (caso não esteja instalado), digitando no terminal: /bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/master/install.sh)"
      - executar no terminal: brew install graphviz
   
 * pacotes python
   - graphviz
   - xlrd
   - odfpy
   - pandas
   - para instalar no Windows
      - executar via linha de comando: "pip install graphviz xlrd odfpy pandas"
   - para instalar no macOS:
      - executar no terminal: "pip3 install graphviz xlrd odfpy pandas"
