# BIM

Ferramentas utilitárias para fluxos BIM.

## Conversor de Excel para Navisworks

O script `navisworks_converter.py` transforma uma planilha Excel (`.xlsx`) em
arquivos XML e XSD compatíveis com o Autodesk Navisworks, preservando a ordem
das planilhas, colunas e campos exatamente como definidos na planilha original.

### Pré-requisitos

- Python 3.10 ou superior
- Biblioteca [openpyxl](https://openpyxl.readthedocs.io/) instalada:

  ```bash
  pip install openpyxl
  ```

### Uso

```bash
python navisworks_converter.py caminho/para/arquivo.xlsx --output diretorio/de/saida
```

Se a opção `--output` não for informada, os arquivos serão gravados no diretório
atual. O script gera dois arquivos com o mesmo nome da planilha de origem, um
com extensão `.xml` e outro `.xsd`.
