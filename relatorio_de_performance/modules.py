def tratar_nome(lista):
    lista = lista.split("-")[-1].replace(".pdf", "")
    return lista

def enable_download(driver, path):
    import os
    driver.command_executor._commands["send_command"] = ("POST", '/session/$sessionId/chromium/send_command')
    params = {'cmd': 'Page.setDownloadBehavior', 'params': {'behavior': 'allow','downloadPath':os.getcwd()+'\%s' % path}}
    driver.execute("send_command", params)

def return_downloaded(nome_assessor):
    from os.path import isfile, join
    from os import listdir
    baixados  = [f for f in listdir("./%s" % nome_assessor) if isfile(join("./%s" % nome_assessor, f))]
    baixados = [tratar_nome(conta) for conta in baixados]
    return baixados