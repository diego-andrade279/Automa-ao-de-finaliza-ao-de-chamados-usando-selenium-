bt_classificacao = "react-select-2-input"
valor = bt_classificacao[13:-6]
print(valor)
bt_classificacao = f"react-select-{int(valor) + 1}-input"
print(bt_classificacao)