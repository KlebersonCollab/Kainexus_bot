#======================================================#
#   Bot para extração de relatório do Sistema Kainexus
#   Autor: Kleberson Romero
#   Data: 09/05/2022
#======================================================#
require 'watir'
require 'inifile'
#Carrega as configurações
@load = IniFile.load('config.ini')
@config = @load['Config']
#Cria um Web Application
@browser = Watir::Browser.new :firefox
@browser.goto @config["Site"]
#Busca o Textinput do usuario Microsoft e seta o valor do usuário do arquivo de configurações
user = @browser.text_field(id: 'i0116')
user.value = @config["User"]
#clica em NextButton
next_button = @browser.element(id: 'idSIButton9')
next_button.click
#Aguarda o carregamento da página
sleep(5)
#Seta o Password do usuario do arquivo de configurações e clica para fazer login
password = @browser.text_field(id: 'i0118')
password.value = @config["Pass"]
#sleep(5)
login_button = @browser.element(id: 'idSIButton9')
login_button.click
#Aguarda o carregamento da página
sleep(10)
yes = @browser.element(id: 'idSIButton9')
yes.click
#Aguarda o carregamento da página do Kainexus
sleep(5)
#Clica na opção itens
tab_itens = @browser.element(id: 'app-idea-list-filter-picker-1040-btnInnerEl')
tab_itens.click
#Clicka em All Itens
all_itens = @browser.element(id: 'menuitem-1083-itemEl')
all_itens.click
#Aguarda o carregamento da página
sleep(7)
#Clica em Actions
actions = @browser.element(id: 'button-1774-btnInnerEl')
actions.click
#Clica em Exportar
export = @browser.element(id: 'menuitem-1746-textEl')
export.text == 'Export'
export.click
#Clica em exportar XLSX
xls = @browser.element(class: 'app-dotted-underline')
xls.text == 'Salvar XLSX'
xls.click
#Clica em Salvar
save_button = @browser.element(id: 'button-2236-btnInnerEl')
save_button.click
sleep(10)
@browser.close
