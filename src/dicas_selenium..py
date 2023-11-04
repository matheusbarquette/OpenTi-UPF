from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By

#--------------------------------------------------------------------------------------------------------------------------------
# WebDriver: O ponto de partida do Selenium é instanciar um objeto WebDriver, que permite interagir com o navegador. Por exemplo:
# Inicialização do WebDriver para o navegador Chrome
# link chromedriver https://chromedriver.chromium.org/downloads, lembrando sempre de verificar a versao do navegador
# OBS: versoes novas nao e necessario
caminho_chromedriver = '/chromedriver.exe'
driver = webdriver.Chrome(caminho_chromedriver)

driver = webdriver.Chrome()
#--------------------------------------------------------------------------------------------------------------------------------


#--------------------------------------------------------------------------------------------------------------------------------
# Navegação: O Selenium permite navegar para diferentes URLs, clicar em links e navegar para frente ou para trás no histórico do navegador. Alguns métodos úteis são:

# Navegar para uma URL específica
driver.get("https://www.example.com")

# Clicar em um link
driver.find_element(By.ID, 'element-id').click()

# Voltar no histórico do navegador
driver.back()

# Avançar no histórico do navegador
driver.forward()
#--------------------------------------------------------------------------------------------------------------------------------


#--------------------------------------------------------------------------------------------------------------------------------
# Interagindo com elementos: Uma vez que você localizou um elemento, o Selenium permite interagir com ele, preenchendo campos de entrada, clicando em botões, selecionando opções em menus suspensos, entre outros. Alguns métodos úteis são:

# Preencher um campo de entrada
driver.find_element(By.ID, 'element-id').send_keys("Texto de exemplo")

# Selecionar uma opção em um menu suspenso (selecionando pelo valor)
select = Select(driver.find_element(By.ID, 'select-element-id'))
select.select_by_value("option-value")
#--------------------------------------------------------------------------------------------------------------------------------


#--------------------------------------------------------------------------------------------------------------------------------
# Esperas: O Selenium também oferece recursos para lidar com esperas, permitindo que você aguarde a ocorrência de determinados eventos antes de prosseguir. Por exemplo:

# Esperar até que um elemento esteja visível
element = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.ID, "element-id")))
#--------------------------------------------------------------------------------------------------------------------------------


#--------------------------------------------------------------------------------------------------------------------------------
# Localizacao de elementos, ID, CLASS E XPATH

# Localizar por ID:  é usado para localizar elementos com base no valor do atributo id do elemento HTML. O ID é único dentro de uma página, o que significa que cada elemento deve ter um ID exclusivo. Portanto, a localização por ID é uma maneira rápida e eficiente de encontrar um elemento específico se ele tiver um ID atribuído. Exemplo:

driver.find_element(By.ID, 'element-id')

# Localizar por classe: é usado para localizar elementos com base no valor do atributo class do elemento HTML. Uma classe pode ser atribuída a vários elementos na página e um elemento pode ter várias classes. Portanto, a localização por classe pode retornar uma lista de elementos correspondentes. Você pode iterar sobre a lista ou usar um índice específico para acessar o elemento desejado. Exemplo:

driver.find_element(By.CLASS_NAME, 'element-class')

# Localizar por XPath: O XPath é uma linguagem de consulta usada para localizar elementos em um documento XML ou HTML. No Selenium, você pode usar para localizar elementos com base em um XPath específico. O XPath fornece uma maneira mais flexível de localizar elementos, permitindo a seleção com base em vários critérios, como a hierarquia de elementos, atributos, texto e muito mais. Exemplo:

driver.find_element(By.XPATH, "//input[@id='element-id']")

# A escolha de qual método usar depende das características específicas do elemento que você está procurando e da estrutura do código HTML da página. Se um elemento tiver um ID exclusivo, a localização por ID é geralmente a opção mais rápida e recomendada. Se você estiver lidando com vários elementos que possuem a mesma classe, a localização por classe é útil. Já o XPath é mais poderoso e pode ser usado para realizar seleções complexas de elementos com base em diferentes critérios.

# É importante observar que existem outros métodos disponíveis para localizar elementos no Selenium, como localização por nome de tag, nome de atributo, texto de link, entre outros. A escolha do método correto depende da estrutura da página e dos requisitos da automação que você está construindo.
#--------------------------------------------------------------------------------------------------------------------------------
