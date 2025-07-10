# 1. Carregar Pacotes Necessários ------------------------------------------------
# Certifique-se de que estão instalados: install.packages(c("ggplot2", "dplyr", "openxlsx"))
library(ggplot2)
library(dplyr)
library(openxlsx)

# 2. Criar o Dataframe a partir dos Dados da Imagem -----------------------------
# Corrigido com os nomes corretos das colunas e valores com ponto decimal.
dados_sifilis <- data.frame(
  ANO = 2009:2021,
  TAXA_SIFILIS_GESTANTE = c(2.13, 2.46, 3.65, 4.56, 6.47, 8.54, 9.53, 12.49, 13.75, 19.32, 18.05, 19.04, 23.53),
  TAXA_SIFILIS_CONGENITA = c(1.21, 1.51, 2.12, 2.70, 4.07, 4.54, 5.67, 6.95, 6.74, 7.51, 6.33, 6.97, 7.88),
  CASOS_SIFILIS_GESTANTE = c(464, 523, 787, 958, 1317, 1747, 1978, 2497, 2811, 3970, 3565, 3599, 4372),
  CASOS_SIFILIS_CONGENITA = c(265, 321, 457, 568, 828, 928, 1177, 1390, 1379, 1543, 1251, 1318, 1464)
)

# 3. Definir Estilo Profissional (Cores e Tema) ----------------------------------
cores_gestante <- list(barra = "#f7a278", linha = "#c94c4c", texto = "#8c2a2a")
cores_congenita <- list(barra = "#68a2b9", linha = "#175676", texto = "#0f3a50")

tema_profissional <- function() {
  theme_minimal(base_family = "sans") +
    theme(
      plot.background = element_rect(fill = "#f5f5f5", color = NA),
      panel.background = element_rect(fill = "#f5f5f5", color = NA),
      panel.grid.major = element_line(color = "white", linewidth = 0.5),
      panel.grid.minor = element_blank(),
      plot.title = element_text(face = "bold", size = 16, color = "#333333", hjust = 0.5),
      plot.subtitle = element_text(size = 12, color = "#555555", hjust = 0.5, margin = margin(b = 15)),
      axis.title = element_text(face = "bold", color = "#555555"),
      axis.text = element_text(color = "#333333"),
      axis.text.x = element_text(angle = 45, hjust = 1)
    )
}

# 4. Criar os dois objetos de gráfico (com nomes de variáveis verificados) -----

# --- Gráfico 1: Sífilis em Gestantes ---
fator_escala_gestante <- max(dados_sifilis$CASOS_SIFILIS_GESTANTE) / max(dados_sifilis$TAXA_SIFILIS_GESTANTE)
grafico_gestante <- ggplot(dados_sifilis, aes(x = as.factor(ANO))) +
  geom_col(aes(y = CASOS_SIFILIS_GESTANTE), fill = cores_gestante$barra, alpha = 0.9) +
  geom_line(aes(y = TAXA_SIFILIS_GESTANTE * fator_escala_gestante, group = 1), color = cores_gestante$linha, linewidth = 1.2) +
  geom_point(aes(y = TAXA_SIFILIS_GESTANTE * fator_escala_gestante), color = cores_gestante$linha, size = 2.5) +
  geom_text(aes(y = TAXA_SIFILIS_GESTANTE * fator_escala_gestante, label = format(TAXA_SIFILIS_GESTANTE, nsmall = 2)), vjust = -1, color = cores_gestante$texto, size = 3.0, fontface = "bold") +
  scale_y_continuous(name = "Número de Casos", sec.axis = sec_axis(~ . / fator_escala_gestante, name = "Taxa de Detecção (por 1.000 nascidos vivos)")) +
  labs(title = "Evolução da Sífilis em Gestantes - Bahia", subtitle = "Número de Casos e Taxa de Detecção (2009-2021)", x = "Ano", y = "Número de Casos") +
  tema_profissional()

# --- Gráfico 2: Sífilis Congênita (COM O ERRO DE DIGITAÇÃO CORRIGIDO) ---
fator_escala_congenita <- max(dados_sifilis$CASOS_SIFILIS_CONGENITA) / max(dados_sifilis$TAXA_SIFILIS_CONGENITA)
grafico_congenita <- ggplot(dados_sifilis, aes(x = as.factor(ANO))) +
  geom_col(aes(y = CASOS_SIFILIS_CONGENITA), fill = cores_congenita$barra, alpha = 0.9) +
  geom_line(aes(y = TAXA_SIFILIS_CONGENITA * fator_escala_congenita, group = 1), color = cores_congenita$linha, linewidth = 1.2) +
  geom_point(aes(y = TAXA_SIFILIS_CONGENITA * fator_escala_congenita), color = cores_congenita$linha, size = 2.5) +
  geom_text(aes(y = TAXA_SIFILIS_CONGENITA * fator_escala_congenita, label = format(TAXA_SIFILIS_CONGENITA, nsmall = 2)), vjust = -1, color = cores_congenita$texto, size = 3.0, fontface = "bold") +
  scale_y_continuous(name = "Número de Casos", sec.axis = sec_axis(~ . / fator_escala_congenita, name = "Taxa de Incidência (por 1.000 nascidos vivos)")) +
  labs(title = "Evolução da Sífilis Congênita - Bahia", subtitle = "Número de Casos e Taxa de Incidência (2009-2023)", x = "Ano", y = "Número de Casos") +
  tema_profissional()

# 5. Exportar para um único Arquivo Excel com duas abas -------------------------
output_file <- "Relatorio_Sifilis_Bahia_Corrigido.xlsx"
wb <- createWorkbook()

# --- Aba 1: Sífilis em Gestante ---
addWorksheet(wb, "Sifilis_Gestante")
writeData(wb, "Sifilis_Gestante", dados_sifilis)
grafico_gestante_path <- "grafico_gestante.png"
ggsave(grafico_gestante_path, plot = grafico_gestante, width = 11, height = 7, dpi = 300)
insertImage(wb, "Sifilis_Gestante", grafico_gestante_path, startRow = 2, startCol = 7, width = 11, height = 7)

# --- Aba 2: Sífilis Congênita ---
addWorksheet(wb, "Sifilis_Congenita")
writeData(wb, "Sifilis_Congenita", dados_sifilis)
grafico_congenita_path <- "grafico_congenita.png"
ggsave(grafico_congenita_path, plot = grafico_congenita, width = 11, height = 7, dpi = 300)
insertImage(wb, "Sifilis_Congenita", grafico_congenita_path, startRow = 2, startCol = 7, width = 11, height = 7)

# --- Salvar o arquivo final do Excel ---
saveWorkbook(wb, output_file, overwrite = TRUE)

# --- Limpar arquivos de imagem temporários ---
file.remove(grafico_gestante_path, grafico_congenita_path)

message("✅ Relatório em Excel '", output_file, "' foi criado com sucesso, contendo as duas abas com os respectivos gráficos.")

