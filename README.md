# nexcafe-vscyber
Converte BD Nexcafe para o VSCyber

## Pré-requisitos:
* Python (testei na versão 3.7.4)
* Firebird instalado na porta padrão (3050)

## Passo a passo:
1. Clone ou download [nexcafe-vscyber](https://github.com/renatovbvargas/nexcafe-vscyber)
2. Exportar os dados do Nexcafe com nome Exportar.xls ([Ver roteiro aqui!](http://www.vscyber.com/wiki/index.php/NexCafe))
3. Executar o script conforme exemplo abaixo de hora a 2,50 reais:
```
python nexcafe-vscyber.py 2.5
```

## Resultado esperado:
* Arquivo VSCyber.FDB e VSCyber.zip na pasta do script
