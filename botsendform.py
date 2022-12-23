import pandas as pd

tabela = pd.read_excel("EXCEL AQUI")
display(tabela)

for i, nome in enumerate(tabela["Nome"]):
    email = tabela.loc[i, "Email"]
    assunto = tabela.loc[i, "Assunto"]
    mensagem = tabela.loc[i, "Mensagem"]
    
    from selenium import webdriver

    navegador = webdriver.Chrome()
    navegador.get("LINK AQUI")
    
    
    # preencher Nome
    navegador.find_element("xpath", '//*[@id="jform_contact_name"]').send_keys(nome)

    # preencher email
    navegador.find_element("xpath", '//*[@id="jform_contact_email"]').send_keys(email)
    #preencher assunto
    navegador.find_element("xpath", '//*[@id="jform_contact_emailmsg"]').send_keys(assunto)
    #preencher mensagem
    navegador.find_element("xpath", '//*[@id="jform_contact_message"]').send_keys(mensagem)

    #marcar checkbox
    marcar = navegador.find_element("xpath", '//*[@id="jform_consentbox0"]')
    navegador.execute_script("arguments[0].click();", marcar)

    #clicar no botao
    enviar = navegador.find_element("xpath", '//*[@id="contact-form"]/div/div/button')
    navegador.execute_script("arguments[0].click();", enviar)