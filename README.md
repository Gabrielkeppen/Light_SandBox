# Sandbox Gráfico Interativo para Visualização de Iluminação e Conceitos Matemáticos em Computação Gráfica

Projeto A3 da disciplina **Computação Gráfica e Realidade Virtual** (UniCuritiba).
Professor: Rubem Matimoto Koide.

## Integrantes

| Integrante | RA |
|---|---|
| Gabriel Thiago Casini Keppen | 172422353 |
| Felipe Silva Muchinski | 172422440 |

## Sobre o projeto

A computação gráfica usa vetores, matrizes, transformações geométricas e produtos escalares a todo momento para representar, transformar e iluminar objetos 3D. Nos motores gráficos modernos, porém, grande parte dessas operações acontece de forma automática e não é percebida por quem usa a ferramenta.

Este projeto propõe um **sandbox gráfico interativo em 3D, desenvolvido na Unreal Engine**, que torna visível a relação entre os conceitos matemáticos da computação gráfica e o resultado visual na cena. O usuário pode manipular fontes de luz, posição e orientação de objetos e propriedades de materiais, e observar em tempo real como isso afeta a iluminação.

### Interface de visualização matemática

Ao analisar uma superfície, o sistema apresenta valores atualizados em tempo real, como:

```
Normal da superfície   N = (0.00, 0.71, 0.71)
Direção da luz         L = (0.42, 0.63, 0.64)
Produto escalar        N · L = 0.89
Ângulo de incidência   θ = 27,1°
Intensidade difusa     89%
```

Além dos números, os próprios vetores (normal, direção da luz etc.) são desenhados no ambiente 3D, diretamente sobre os objetos analisados.

## Escopo

**Produto mínimo viável (MVP)**
- Ambiente 3D na Unreal Engine
- Fonte de luz manipulável
- Visualização dos vetores (normal e direção da luz)
- Cálculo e exibição da iluminação difusa (N · L, ângulo de incidência, intensidade)

**Módulos futuros (estrutura modular)**
- Reflexão especular e vetores de reflexão
- Transformações de posição, rotação e escala
- Matrizes de transformação
- Propriedades de materiais
- Diferentes tipos de fontes de luz
- Comportamento de sombras

## Ferramentas

- **Unreal Engine** (Blueprints e, se necessário, C++): ambiente 3D, iluminação, materiais, renderização em tempo real, interface e interação
- Conceitos de CG explorados: vetores, produto escalar, matrizes, transformações 3D, iluminação, materiais, shaders, sombras e renderização em tempo real

## Cronograma e entregas

| Entrega | Data | Status |
|---|---|---|
| Proposta inicial (conceitual) | 07/09/2026 | Entregue |
| Entrega parcial (documento ABNT: introdução, revisão bibliográfica, desenvolvimento parcial, referências) | 12/10/2026 | Em andamento |
| Avaliação A3 (relatório ABNT e apresentação na Mostra UniCuritiba) | 30/11 a 05/12/2026 | Pendente |

## Estrutura do repositório

Esta pasta guarda tudo o que é necessário para o projeto na Unreal Engine (o projeto `.uproject`, conteúdo, código e documentação). O projeto Unreal ainda será criado nesta pasta.

## Status atual

- [x] Proposta conceitual entregue
- [x] Repositório criado
- [ ] Projeto Unreal criado
- [ ] Cena base com superfície e fonte de luz
- [ ] Blueprint do cálculo N · L e ângulo de incidência
- [ ] Widget de interface com os valores em tempo real
- [ ] Vetores desenhados na cena
