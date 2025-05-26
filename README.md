# SPDO Resumo de Fichas Técnicas

Bem-vindo ao **SPDO Resumo de Fichas Técnicas**, uma aplicação web construída com Streamlit que permite gerar, de forma simples e rápida, um relatório consolidado de metadados a partir de uma pasta de arquivos Excel (.xlsx).

---

## 🎯 Fluxo de uso

1. **Upload do ZIP**

   * Clique em "Browse files" e selecione o **Fichas Tecnicas.zip**.
   * O arquivo deve conter a estrutura:

     ```
     Fichas Tecnicas/<NomeCliente>/<Categoria>/<Arquivo>.xlsx
     ```

2. **Processamento**

   * Clique em **Iniciar Processamento**.
   * Acompanhe a barra de progresso e o status de processamento.

3. **Download do resumo**

   * Após a conclusão, clique em **📥 Baixar Resumo como Excel**.
   * O arquivo gerado (`Resumo_Fichas_Tecnicas.xlsx`) terá uma aba para cada cliente, com as seguintes colunas:

     * `Arquivo`: nome do arquivo original
     * `Tipo`: categoria (Materiais/Equipamentos)
     * `Criação`: data de criação (DD/MM/AAAA)
     * `Atualização`: data de última atualização
     * `Código Externo`: identificador externo

---

```
