# Simulador de Fundos Imobiliários (Excel)

Ferramenta em Excel que simula um investimento mensal em fundos imobiliários: a partir do aporte, do prazo e da taxa de rendimento, ela calcula o patrimônio acumulado, os dividendos mensais e como esse aporte se divide entre seis tipos de FII conforme o perfil de investidor.

Projeto do desafio de Excel da DIO (Expert Felipe Aguiar).

**Arquivo:** [`simulador-fundos-imobiliarios.xlsx`](./simulador-fundos-imobiliarios.xlsx)

> Todos os valores do arquivo são **de exemplo** (salário, aporte, taxa). Os percentuais por perfil também são exemplo, não recomendação de investimento.

---

## As perguntas que a ferramenta responde

| # | Pergunta | Onde aparece |
|---|----------|--------------|
| 1 | Quanto investir por mês? | `Simulador!C7` — entrada |
| 2 | Por quantos anos? | `Simulador!C8` — entrada |
| 3 | Qual a taxa de rendimento mensal? | `Simulador!C9` — entrada |
| 4 | Quanto de patrimônio vou acumular? | `Simulador!C12` — calculado |
| 5 | Quanto vou receber de dividendos por mês? | `Simulador!C13` — calculado |

Quem usa só mexe nas **células amarelas**. As **células cinzas** são calculadas e não devem ser editadas. As linhas de grade estão escondidas e a planilha tem banner, para ter cara de aplicativo.

### O que mais tem na tela

- **Projeção de patrimônio** para 2, 5, 10, 20 e 30 anos (`F7:I12`)
- **Configurações**: salário, % sugerido de aporte (30%) e uma segunda simulação usando esse aporte sugerido (`B18:C23`)
- **Divisão do aporte** entre 6 tipos de FII conforme o perfil escolhido numa lista (`B26:D34`)
- **Gráfico de pizza** da divisão, que muda junto com o perfil
- Aba **Apoio** com a tabela de perfis e a conferência de que cada perfil soma 100%

---

## Como o VF e o PROCV entram nos cálculos

### VF — patrimônio acumulado

```excel
C12 = -VF(taxa_mensal; prazo_anos*12; aporte)
```

- `taxa_mensal` — a taxa de cada período. Como o aporte é mensal, a taxa também é mensal.
- `prazo_anos*12` — o número de períodos. Os anos viram meses para ficar na mesma unidade da taxa.
- `aporte` — o valor depositado a cada período. Entra negativo na lógica do Excel (é dinheiro que sai do bolso), por isso a fórmula leva um **sinal de menos na frente** para o patrimônio aparecer positivo.

A partir do patrimônio vêm os dividendos:

```excel
C13 = patrimonio * taxa_mensal
```

Os cenários repetem o mesmo VF trocando só o prazo, lendo o número de anos da própria linha:

```excel
H8 = -VF(taxa_mensal; $F8*12; aporte)
```

O `$F8` trava a coluna para a fórmula não se perder ao arrastar, e `taxa_mensal` e `aporte` são intervalos nomeados, que já são absolutos por natureza.

### PROCV com chave composta — divisão por perfil

O percentual depende de **duas** informações ao mesmo tempo (o perfil **e** o tipo de fundo), e o PROCV procura por uma coluna só. A saída é juntar as duas numa **chave composta**.

Na aba `Apoio`, a coluna A monta a chave de cada linha:

```excel
A3 = B3 & "-" & C3     ->  "Conservador-Lajes Corporativas"
```

Na aba `Simulador`, a busca monta a mesma chave com o perfil escolhido na lista:

```excel
C28 = PROCV(perfil & "-" & $B28; tab_perfis; 4; FALSO)
D28 = aporte * C28
```

Trocar o perfil na lista muda a chave, o PROCV acha outra linha e a divisão inteira (e o gráfico) se refaz.

A lista do perfil é uma **validação de dados** apontando para `Apoio!F3:F5`.

---

## Intervalos nomeados criados

| Nome | Referência | O que é |
|------|-----------|---------|
| `aporte` | `Simulador!$C$7` | Aporte mensal |
| `prazo_anos` | `Simulador!$C$8` | Prazo em anos |
| `taxa_mensal` | `Simulador!$C$9` | Taxa de rendimento mensal |
| `patrimonio` | `Simulador!$C$12` | Patrimônio acumulado |
| `salario` | `Simulador!$C$18` | Salário mensal |
| `pct_sugerido` | `Simulador!$C$19` | % do salário sugerido para aporte |
| `aporte_sugerido` | `Simulador!$C$20` | Salário × percentual sugerido |
| `perfil` | `Simulador!$C$26` | Perfil escolhido na lista |
| `tab_perfis` | `Apoio!$A$3:$D$20` | Tabela de busca do PROCV |

Usar nomes no lugar de `$C$7` deixa a fórmula legível: `=-VF(taxa_mensal; prazo_anos*12; aporte)` se lê quase como uma frase.

---

## Percentuais de cada perfil

| Tipo de fundo | Conservador | Moderado | Arrojado |
|---|---|---|---|
| Lajes Corporativas | 10% | 15% | 15% |
| Shoppings | 5% | 15% | 25% |
| Galpões Logísticos | 15% | 20% | 25% |
| Papel (CRI) | 40% | 25% | 10% |
| Fundo de Fundos (FOF) | 20% | 15% | 15% |
| Híbrido | 10% | 10% | 10% |
| **Total** | **100%** | **100%** | **100%** |

**De onde vieram:** são percentuais que eu montei para o exercício, seguindo a lógica de que fundos de papel (CRI) são mais previsíveis e pesam mais no perfil conservador, enquanto shoppings e galpões, mais ligados ao ciclo econômico, pesam mais no arrojado. Não são recomendação de investimento. A aba `Apoio` tem uma conferência com SOMASE mostrando que cada perfil fecha em 100%.

---

## O que mudei em relação à ferramenta do Expert

- **Banner e paleta próprios**, em vez do banner da aula.
- **Percentuais próprios** por perfil, com uma conferência automática (SOMASE) de que cada um soma 100%.
- **Duas linhas a mais no resultado**: total aportado no período e rendimento gerado, para separar o que foi dinheiro meu do que foi juros composto.
- **Segunda simulação** rodando em paralelo com o aporte sugerido de 30% do salário, ao lado da simulação do aporte informado.
- **Total aportado também na tabela de cenários**, para comparar esforço e resultado em cada prazo.
- **Gráfico de pizza** da divisão do aporte, ligado às células do PROCV — muda sozinho ao trocar o perfil.
- **Rendimento anual equivalente** da carteira, convertendo a taxa mensal: `(1+taxa_mensal)^12-1`.

---

## Checklist de entrega

- [x] A planilha abre sem erro de fórmula e responde às cinco perguntas
- [x] Trocar o perfil na lista muda a divisão do aporte
- [x] Cada perfil soma 100%
- [x] Salário e demais valores publicados são de exemplo
- [x] Repositório público, nome em minúsculas e sem acento
- [ ] Prints da ferramenta preenchida nos dois perfis (adicionar em `/prints`)

---

## Tecnologias

Microsoft Excel — funções `VF`, `PROCV`, `SOMASE`; intervalos nomeados; validação de dados; formatação condicional de tela (grade oculta, células calculadas em cinza); gráfico de pizza.
