# 🧠 Belmicro – Automação e Análise de Preços (Python)

Pipeline completo para **coletar**, **limpar** e **comparar** preços da Belmicro vs concorrentes na Shopee, gerando **sugestões de preço** (estratégia: Belmicro em 3º menor preço).

## 🚀 Stack
- Python 3.x · Pandas · Selenium/undetected-chromedriver · Playwright (tests) · openpyxl
- (Opcional) Groq API para apoio na deduplicação/normalização de produtos

## 🗂 Estrutura

## 🧠 Lógica do Pipeline
1. **Coleta** – Selenium navega na Shopee e extrai nome, vendedor e preço.  
2. **Limpeza** – Pandas padroniza nomes e remove duplicatas.  
3. **Análise** – Compara produtos e define preço sugerido.  
4. **Saída** – Planilha final pronta para análise de pricing.

## ▶️ Como Executar
```bash
python -m venv .venv
.\.venv\Scripts\activate
pip install -r requirements.txt

python 1_coleta_bruta/robo_coleta.py
python 2_limpeza/limpeza_planilha.py
python 3_sugestao_preco/sugestao_preco.py
