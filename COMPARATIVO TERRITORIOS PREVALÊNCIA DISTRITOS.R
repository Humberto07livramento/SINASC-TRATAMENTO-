# Instalar e carregar os pacotes necessários
if (!require("readxl"))   install.packages("readxl",   dependencies = TRUE)
if (!require("dplyr"))    install.packages("dplyr",    dependencies = TRUE)
if (!require("tidyr"))    install.packages("tidyr",    dependencies = TRUE)
if (!require("ggplot2"))  install.packages("ggplot2",  dependencies = TRUE)
if (!require("openxlsx")) install.packages("openxlsx", dependencies = TRUE)

library(readxl)
library(dplyr)
library(tidyr)
library(ggplot2)
library(openxlsx)

# --- ALTERAÇÃO AQUI ---
# Definir o diretório de saída explicitamente conforme solicitado
doc_dir <- "C:/Users/humbertolivramento/Documents"

# Função para ler e preparar cada ano
# Assumindo que os arquivos de entrada ("anomalia_2011.xlsx", "anomalia_2024.xlsx") 
# também estão neste mesmo diretório.
envelope_ano <- function(file, ano) {
  df <- read_xlsx(file.path(doc_dir, file))
  names(df)[1] <- "Indicador"
  df <- df %>% filter(!is.na(Indicador))
  df_long <- df %>% pivot_longer(-Indicador, names_to = "Municipio", values_to = "Valor")
  df_wide <- df_long %>% pivot_wider(names_from = Indicador, values_from = Valor)
  names(df_wide) <- tolower(names(df_wide))
  names(df_wide) <- gsub("-", "_", names(df_wide))
  df_wide <- df_wide %>%
    rename(
      AC_Sim        = ac_sim,
      AC_Nao        = ac_nao,
      AC_Ignorado   = ac_ignorado,
      NascidosVivos = nascidosvivos,
      Municipio     = municipio
    ) %>%
    mutate(Ano = ano)
  return(df_wide)
}

# Preparar dados
dados_2011 <- envelope_ano("anomalia_2011.xlsx", 2011)
dados_2024 <- envelope_ano("anomalia_2024.xlsx", 2024)
dados       <- bind_rows(dados_2011, dados_2024)

# Cálculos
dados <- dados %>%
  mutate(
    Denom_Ajustado = NascidosVivos - AC_Ignorado,
    Prevalencia    = round((AC_Sim / Denom_Ajustado) * 10000, 2),
    se             = sqrt((AC_Sim / Denom_Ajustado^2) * 10000),
    IC_LI          = round(Prevalencia - 1.96 * se, 2),
    IC_LS          = round(Prevalencia + 1.96 * se, 2)
  ) %>%
  select(Municipio, Ano, NascidosVivos, AC_Sim, AC_Nao, AC_Ignorado,
         Denom_Ajustado, Prevalencia, IC_LI, IC_LS)

# Exportar tabelas para Excel
dados_com_ic_2011 <- dados %>% filter(Ano == 2011)
dados_com_ic_2024 <- dados %>% filter(Ano == 2024)

tabelas_para_exportar <- list(
  "Resultados_2011" = dados_com_ic_2011,
  "Resultados_2024" = dados_com_ic_2024
)

nome_arquivo_excel <- "prevalencia_anomalias_congenitas.xlsx"
write.xlsx(tabelas_para_exportar, file = file.path(doc_dir, nome_arquivo_excel))


# ——————
# Gerar duas figuras com seis distritos cada e aumentar tamanho de fonte

dist_list <- unique(dados$Municipio)
group1 <- dist_list[1:6]
group2 <- dist_list[7:12]

gerar_figura <- function(grupo, nome_arquivo) {
  df_sub <- dados %>% filter(Municipio %in% grupo)
  p <- ggplot(df_sub, aes(x = factor(Ano), y = Prevalencia, color = Municipio, group = Municipio)) +
    geom_line(linewidth = 1) +
    geom_point(size = 3) +
    geom_errorbar(aes(ymin = IC_LI, ymax = IC_LS), width = 0.2, linewidth = 0.7) +
    facet_wrap(~ Municipio, ncol = 3, scales = "fixed", drop = FALSE) +
    theme_minimal() +
    labs(
      title = "Prevalência de AC por Distrito Sanitário",
      subtitle = paste0(min(dados$Ano), " e ", max(dados$Ano)),
      x     = "Ano",
      y     = "Prevalência (por 10.000 NV)",
      color = NULL
    ) +
    theme(
      plot.title      = element_text(face = "bold", size = 18, hjust = 0.5),
      plot.subtitle   = element_text(size = 14, hjust = 0.5),
      axis.title      = element_text(size = 12),
      axis.text.x     = element_text(size = 10, angle = 45, hjust = 1),
      axis.text.y     = element_text(size = 10),
      strip.text      = element_text(size = 12),
      legend.position = "none",
      panel.spacing   = unit(1, "lines")
    )
  
  ggsave(
    filename = file.path(doc_dir, nome_arquivo),
    plot     = p,
    width    = 12,
    height   = 8,
    dpi      = 300
  )
}

# Gerar figuras
gerar_figura(group1, "figura_distritos_1.png")
gerar_figura(group2, "figura_distritos_2.png")

