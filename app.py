import openpyxl
import pyperclip
import pyautogui
from time import sleep

workbook = openpyxl.load_workbook('produtos_nao_cadastrados.xlsx')
sheet_produtos = workbook['Produtos']

for linha in sheet_produtos.iter_rows(min_row=2):
    
    # sessao I
    
    nome_produto = linha[0].value
    pyperclip.copy(nome_produto)
    pyautogui.click(966,431, duration=0.1)
    pyautogui.hotkey('ctrl','v')

    descricao = linha[1].value
    pyperclip.copy(descricao)
    pyautogui.click(967,581, duration=0.1)
    pyautogui.hotkey('ctrl','v')

    categoria = linha[2].value
    pyperclip.copy(categoria)
    pyautogui.click(971,723, duration=0.1)
    pyautogui.hotkey('ctrl','v')

    pyautogui.click(861,798, duration=0.1)
    sleep(0.8)

    # sessao II

    codigo_produto = linha[3].value
    pyperclip.copy(codigo_produto)
    pyautogui.click(905,475, duration=0.1)
    pyautogui.hotkey('ctrl','v')
    
    peso = linha[4].value
    pyperclip.copy(peso)
    pyautogui.click(902,619, duration=0.1)
    pyautogui.hotkey('ctrl','v')

    dimensoes = linha[5].value
    pyperclip.copy(dimensoes)
    pyautogui.click(912,767, duration=0.1)
    pyautogui.hotkey('ctrl','v')

    pyautogui.click(929,848, duration=0.1)
    sleep(0.8)

    # sessao III

    preco = linha[6].value
    pyperclip.copy(preco)
    pyautogui.click(876,474, duration=0.1)
    pyautogui.hotkey('ctrl','v')

    quantidade_em_estoque = linha[7].value
    pyperclip.copy(quantidade_em_estoque)
    pyautogui.click(872,618, duration=0.1)
    pyautogui.hotkey('ctrl','v')

    data_de_validade = linha[8].value
    pyperclip.copy(data_de_validade)
    pyautogui.click(918,771, duration=0.1)
    pyautogui.hotkey('ctrl','v')

    pyautogui.click(954,851, duration=0.1)
    sleep(0.8)

    # sessao IV

    cor = linha[9].value
    pyperclip.copy(cor)
    pyautogui.click(912,469, duration=0.1)
    pyautogui.hotkey('ctrl','v')

    tamanho = linha[10].value
    pyperclip.copy(tamanho)
    pyautogui.click(912,625, duration=0.1)
    pyautogui.hotkey('ctrl','v')   

    material = linha[11].value
    pyperclip.copy(material)
    pyautogui.click(911,767, duration=0.1)
    pyautogui.hotkey('ctrl','v')

    pyautogui.click(940,839, duration=0.1)
    sleep(0.8)

    # sessão V

    fabricante = linha[12].value
    pyperclip.copy(fabricante)
    pyautogui.click(921,471, duration=0.1)
    pyautogui.hotkey('ctrl','v')

    pais_origem = linha[13].value
    pyperclip.copy(pais_origem)
    pyautogui.click(916,622, duration=0.1)
    pyautogui.hotkey('ctrl','v') 

    observacoes = linha[14].value
    pyperclip.copy(observacoes)
    pyautogui.click(910,771, duration=0.1)
    pyautogui.hotkey('ctrl','v')

    pyautogui.click(939,840, duration=0.1)
    sleep(0.8)

    # sessao VI

    codigo_de_barras = linha[15].value
    pyperclip.copy(codigo_de_barras)
    pyautogui.click(896,470, duration=0.1)
    pyautogui.hotkey('ctrl','v')

    localizacao_armazem = linha[16].value
    pyperclip.copy(localizacao_armazem)
    pyautogui.click(901,620, duration=0.1)
    pyautogui.hotkey('ctrl','v') 

    pyautogui.click(935,690, duration=0.1)
    sleep(0.8)

    # Enviar outra Resposta

    pyautogui.click(898,267, duration=0.1)
    sleep(0.8)