import pandas as pd
import os

caminho_ficheiro = "Transporte Gestores (1).xlsx"

print("\n" + "="*50)
print(f"A procurar o ficheiro em: {os.path.abspath(caminho_ficheiro)}")
print("="*50)

try:
    # Tenta ler a folha 'BASE' forçando tudo como texto
    df = pd.read_excel(caminho_ficheiro, sheet_name='BASE', dtype=str)
    
    print("\n✅ FICHEIRO LIDO COM SUCESSO!")
    print("\n👉 COLUNAS QUE O PYTHON CONSEGUE VER:")
    colunas = df.columns.tolist()
    for i, col in enumerate(colunas):
        print(f"  Coluna {i}: '{col}'")
        
    print("\n👉 DADOS DA PRIMEIRA PESSOA (Linha 2 do Excel):")
    primeira_linha = df.iloc[0].to_dict()
    for chave, valor in primeira_linha.items():
        print(f"  {chave}: {valor}")

except Exception as e:
    print(f"\n❌ ERRO AO LER O FICHEIRO: {e}")

print("\n" + "="*50)