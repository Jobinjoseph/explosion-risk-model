# install.packages("Microsoft365R")
# install.packages("openxlsx")
library(httr)
library(Microsoft365R)
library(readxl)
library(openxlsx)
library(tools)
library(ggplot2)
library(dplyr)
#---- Direkt zugang zum Sharepoint folder----
site <- sharepoint_site("xxxxx")
drive <- site$get_drive()
drive$list_items()
items <- drive$list_items("xxxxx")
print(items)
drive$download_file(
  item_id = "xxxxx",
  dest = "results.xlsx"
)
url<-"xxxxx/xx/xxx"
GET(url, write_disk("results.xlsx", overwrite = TRUE))

#---- Direkt zugang zur Lokalen OneDrive----
setwd("xxxx/")

# Source the functions 
srcDir <- paste0(getwd(),"/")      # aktueller Ordner
file <- "WA_Lager_functions.R"
source(paste(srcDir,file, sep=""))


#---- path input Excel file ----
file_path <- paste0(srcDir,"Haupt_Liste_Kammer_xxx.xlsx")
sheet_name_1 <- "Kammern"
sheet_name_2 <- "Stollen"
sheet_name_3 <- "Exposition"

#---- path output Excel file ----
output_file_path <- paste0(srcDir,"ergebnisse")
output_file <- paste0(output_file_path,"/ergebnis_1.xlsx")

#---- Read input data ----
Kammern <- read_explosion_data(file_path, sheet_name_1)
Stollen <- read_explosion_data(file_path, sheet_name_2)
Exposition <- read_explosion_data(file_path, sheet_name_3)
View(Exposition)
#---- Loop durch jede Kammer-Daten ----
if (!dir.exists(output_file)) {
  dir.create(output_file, recursive = TRUE)
}

unique_kammer_values <- unique(Kammern$Kammer)

for (kammer in unique_kammer_values) {
  Kammern <- dk_do(Kammern)
  Kammern <- Po_Tipo(Kammern)
  
}
View(Kammern)
View(Stollen)
#---- Berechnung d (äquivalenter Stollendurchmesser in m) für jede Stollen abschnitt # Ref: TLM Luftstoss Zugangsstollen A3-5, Figur A3-1 ----

Stollen$ds<-sqrt((4 * Stollen$Fs) / pi)


#---- Berechnung LR (Für die Reibung massgebende Stollenlänge in m) für jede Stollen abschnitt # Ref: TLM Luftstoss Zugangsstollen A3-7, Figur A3-3 ----

Stollen<-LR(Stollen)
Stollen$Lr

#---- Berechnung alpha_rau (Wandrauhigkeitsbeiwert) für jede Stollen abschnitt # Ref: TLM Luftstoss Zugangsstollen A3-8, Figur A3-4 ----
# 
Stollen<-alpha_rau_Stollen(Stollen)
Stollen$alpha_rau
str(Stollen)
#---- Berechnung Chi (Abstandkoeffizient) für jede Stollen abschnitt # Ref: TLM Luftstoss Zugangsstollen A3-8, Figur A3-4 ----

Stollen$chi<-Stollen$alpha_rau*Stollen$Lr
Stollen$chi
#---- Po und Tipo von Kammern ins Stollen einfügen ----
Stollen <- Stollen%>%
  left_join(Kammern %>% select(Kammer, Po, Tipo, Q), by = "Kammer")
# Stollen$tau<-Stollen$alpha_rau*Stollen$Tipo

View(Stollen)
#---- Prototyp ganze Code ----
final_results <- data.frame()

for (kammer in unique_kammer_values) {
  
  match_index <- which(Kammern$Kammer == kammer)
  kammer_Stollen <- Stollen[Stollen[[1]] %in% kammer, ]
  p <- kammer_Stollen$Po
  t <- kammer_Stollen$Tipo
  F1 <- numeric(nrow(kammer_Stollen))
  F2 <- numeric(nrow(kammer_Stollen))
  Lo <- numeric(nrow(kammer_Stollen))
  Lss <- numeric(nrow(kammer_Stollen))
  tau <- numeric(nrow(kammer_Stollen))
  LR <- numeric(nrow(kammer_Stollen))
  ds <- numeric(nrow(kammer_Stollen))
  
  for (i in 1:nrow(kammer_Stollen)) {
    print(paste("Processing row:", i))
    
    P_out <- kammer_Stollen$P_out[i]
    chi <- kammer_Stollen$chi[i]
    LR <- kammer_Stollen$Lr[i]
    alpha_rau <- kammer_Stollen$alpha_rau[i]
    alpha_wi <- kammer_Stollen$alpha_wi[i]
    Element <- kammer_Stollen$Element[i]
    Lo <- kammer_Stollen$Lo[i]
    Lss <- kammer_Stollen$Lss[i]
    F1 <- kammer_Stollen$Fs[i]
    F2 <- ifelse(i < nrow(kammer_Stollen), kammer_Stollen$Fs[i+1], NA)
    ds <- kammer_Stollen$ds[i]
    
    
    if (Element == "Verz_1") {
      t <- 0.7 * t_WR(p, t, chi, LR)
      kammer_Stollen<-tau(kammer_Stollen)
      p <- p_Verz_1(P_out, p, t, chi, LR, alpha_rau, alpha_wi)
      
    } else if (Element == "Verz_2") {
      p <- p_Verz_2(P_out, p, t, chi, LR, alpha_rau, alpha_wi)
      t <- 0.7 * t_WR(p, t, chi, LR)
    } else if (Element == "Verz_3") {
      p <- p_Verz_3(P_out, p, t, chi, LR, alpha_rau, alpha_wi)
      t <- 0.7 * t_WR(p, t, chi, LR)
    } else if (Element == "PEW") {
      p <- p_PEW(p, t, F1, F2, chi, LR, alpha_rau)
      t <- t_WR(p, t, chi, LR)
    } else if (Element == "PVE") {
      p <- p_PVE(p, t, F1, F2, chi, LR, alpha_rau)
      t <- t_WR(p, t, chi, LR)
    } else if (Element == "BL") {
      p <- p_BL(p, t, F1, F2, chi, LR, alpha_rau)
      t <- t_WR(p, t, chi, LR)
    }else if (Element == "SM") {
      p <- p_SM( p, t, chi, LR, alpha_rau)
      t <- t_WR(p, t, chi, LR)
    }
    
    p2[i] <- p
    t2[i] <- t
    }
    

  exposition_kammer <- unique(Exposition[Exposition$Kammer == kammer, ])
  if (nrow(exposition_kammer) == 0) {
    cat("No exposition data found for Kammer:", kammer, "\n")
    next
  }
  
  for (j in 1:nrow(exposition_kammer)) {
    Stollenmund <- exposition_kammer$Stollenmund[j]
    Objekt_Nr <- exposition_kammer$Objekt_Nr[j]
    Mund_X <- exposition_kammer$Mund_X[j]
    Mund_Y <- exposition_kammer$Mund_Y[j]
    Objekt_X <- exposition_kammer$Objekt_X[j]
    Objekt_Y <- exposition_kammer$Objekt_Y[j]
    Stollen_X <- exposition_kammer$Stollen_X[j]
    Stollen_Y <- exposition_kammer$Stollen_Y[j]
    d <- sqrt((Objekt_X - Mund_X)^2 + (Objekt_Y - Mund_Y)^2)
    numerator <- (Objekt_X - Mund_X) * (Mund_X - Stollen_X) + (Objekt_Y - Mund_Y) * (Mund_Y - Stollen_Y)
    denominator <- sqrt((Mund_X - Stollen_X)^2 + (Mund_Y - Stollen_Y)^2) * sqrt((Objekt_X - Mund_X)^2 + (Objekt_Y - Mund_Y)^2)
    omega <- acos(numerator / denominator)
    omega_degrees <- omega * (180 / pi)
    alpha <- asin((0.43 / 0.57) * sin(omega))
    alpha_degrees <- alpha * (180 / pi)
    gamma <- 180 - alpha_degrees - omega_degrees
    r <- sqrt(d^2 / (0.43^2 + 0.57^2 - 2 * 0.43 * 0.57 * cos(gamma * pi / 180)))
    p_final <- p / ((r / (0.7 * ds)) / 0.9)
    t_final <- p^(5/3) * t * (0.36 * ds / r)^2
    
    final_results <- rbind(final_results, data.frame(
      Kammer = kammer,
      Stollenmund = Stollenmund,
      Objekt_Nr = Objekt_Nr,
      Stollen_X = Stollen_X,
      Stollen_Y = Stollen_Y,
      Mund_X = Mund_X,
      Mund_Y = Mund_Y,
      Objekt_X = Objekt_X,
      Objekt_Y = Objekt_Y,
      Po = p,
      Tipo = t,
      p_final = p_final,
      t_final = t_final,
      stringsAsFactors = FALSE
    ))
  }
  
  cat("Processed row:", i, "p:", p, "t:", t, "F1:", F1, "F2:", F2, "tau:", tau, "p_final:", p_final, "t_final:", t_final, "\n")
}

if (!dir.exists("ergebnisse")) {
  dir.create("ergebnisse")
}
timestamp <- format(Sys.time(), "%Y%m%d")
file_name <- paste0("ergebnisse/final_results_", timestamp, ".csv")
write.csv(final_results, file = file_name, row.names = FALSE)
cat("Final results saved to", file_name, "\n")


#---- Abschnitt 1: Druck Und Tip Calculation (Stollen Loop)----

if (!dir.exists("ergebnisse")) {
  dir.create("ergebnisse")
}


intermediate_results <- data.frame()

for (kammer in unique_kammer_values) {
  
  match_index <- which(Kammern$Kammer == kammer)
  kammer_Stollen <- Stollen[Stollen[[1]] %in% kammer, ]
  p <- kammer_Stollen$Po[1]  
  t <- kammer_Stollen$Tipo[1]  
  F1 <- NA
  F2 <- NA
  tau <- NA
  ds<- NA
  
  for (i in 1:nrow(kammer_Stollen)) {
    print(paste("Processing row:", i))
    
    P_out <- kammer_Stollen$P_out[i]
    chi <- kammer_Stollen$chi[i]
    LR <- kammer_Stollen$Lr[i]
    alpha_rau <- kammer_Stollen$alpha_rau[i]
    alpha_wi <- kammer_Stollen$alpha_wi[i]
    Element <- kammer_Stollen$Element[i]
    Lo <- kammer_Stollen$Lo[i]
    Lss <- kammer_Stollen$Lss[i]
    F1 <- kammer_Stollen$Fs[i]
    F2 <- ifelse(i < nrow(kammer_Stollen), kammer_Stollen$Fs[i+1], NA)
    ds <- kammer_Stollen$ds[i]

    if (Element == "Verz_1") {
      p <- p_Verz_1(P_out, p, t, chi, LR, alpha_rau, alpha_wi)
      t <- 0.7 * t_WR(p, t, chi, LR)
    } else if (Element == "Verz_2") {
      p <- p_Verz_2(P_out, p, t, chi, LR, alpha_rau, alpha_wi)
      t <- 0.7 * t_WR(p, t, chi, LR)
    } else if (Element == "Verz_3") {
      p <- p_Verz_3(P_out, p, t, chi, LR, alpha_rau, alpha_wi)
      t <- 0.7 * t_WR(p, t, chi, LR)
    } else if (Element == "PEW") {
      p <- p_PEW(p, t, F1, F2, chi, LR, alpha_rau)
      t <- t_WR(p, t, chi, LR)
    } else if (Element == "PVE") {
      p <- p_PVE(p, t, F1, F2, chi, LR, alpha_rau)
      t <- t_WR(p, t, chi, LR)
    } else if (Element == "SM") {
      p <- p_SM(p, t, chi, LR, alpha_rau)
      t <- t_WR(p, t, chi, LR)
    }

    tau <- alpha_rau * t
  }

  intermediate_results <- rbind(intermediate_results, data.frame(
    Kammer = kammer,
    p = p,
    t = t,
    F1 = F1,
    F2 = F2,
    tau = tau,
    ds = ds,
    stringsAsFactors = FALSE
  ))
}

timestamp <- format(Sys.Date(), "%Y%m%d")
intermediate_file <- paste0("ergebnisse/intermediate_results_", timestamp, ".csv")
write.csv(intermediate_results, file = intermediate_file, row.names = FALSE)
cat("Final intermediate results saved to", intermediate_file, "\n")


#---- Abschnitt 1 mit alle parameterwerte speicherung ----
if (!dir.exists("ergebnisse")) {
  dir.create("ergebnisse")
}

intermediate_results <- data.frame()
detailed_intermediate_values <- data.frame()
kammer <-101
for (kammer in unique_kammer_values) {
  
  match_index <- which(Kammern$Kammer == kammer)
  kammer_Stollen <- Stollen[Stollen[[1]] %in% kammer, ]
  p <- kammer_Stollen$Po[1]  
  t <- kammer_Stollen$Tipo[1]  
  F1 <- NA
  F2 <- NA
  tau <- NA
  ds <- NA
  for (i in 1:nrow(kammer_Stollen)) {
    cat("Processing row:", i, "in Kammer:", kammer, "\n")
    
    P_out <- kammer_Stollen$P_out[i]
    chi <- kammer_Stollen$chi[i]
    LR <- kammer_Stollen$Lr[i]
    alpha_rau <- kammer_Stollen$alpha_rau[i]
    alpha_wi <- kammer_Stollen$alpha_wi[i]
    Element <- kammer_Stollen$Element[i]
    Lo <- kammer_Stollen$Lo[i]
    Lss <- kammer_Stollen$Lss[i]
    F1 <- kammer_Stollen$Fs[i]
    F2 <- ifelse(i < nrow(kammer_Stollen), kammer_Stollen$Fs[i+1], NA)
    ds <- kammer_Stollen$ds[i]
    
    cat("Element being processed:", Element, "\n")
    cat("Input values → P_out:", P_out, " | p:", p, " | t:", t, " | chi:", chi,
        " | LR:", LR, " | alpha_rau:", alpha_rau, " | alpha_wi:", alpha_wi, "\n")
    
    p1 <- p  # Save original p before processing
    tau <- alpha_rau * t
    tryCatch({
      if (Element == "Verz_1") {
        p <- p_Verz_1(P_out, p, t,tau, chi, LR, alpha_rau)
        t <- 0.7 * t_WR(p, t, chi, LR)
      } else if (Element == "Verz_2") {
        p <- p_Verz_2(P_out, p, t, chi, LR, alpha_rau, alpha_wi)
        t <- 0.7 * t_WR(p, t, chi, LR)
      } else if (Element == "Verz_3") {
        p <- p_Verz_3(P_out, p, t, chi, LR, alpha_rau, alpha_wi)
        t <- 0.7 * t_WR(p, t, chi, LR)
      }else if (Element == "Stollen") {
        p <- p_WR(p, t, tau, chi, LR, alpha_rau)
        t <- t_WR(p, t, chi, LR)
      } else if (Element == "PEW") {
        p <- p_PEW(p, t, F1, F2, chi, LR, alpha_rau)
        t <- t_WR(p, t, chi, LR)
      } else if (Element == "PVE") {
        p <- p_PVE(p, t, F1, F2, chi, LR, alpha_rau)
        t <- t_WR(p, t, chi, LR)
      } else if (Element == "BL") {
        p <- p_BL(p, t, F1, F2, chi, LR, alpha_rau)
        t <- t_WR(p, t, chi, LR)
      }else if (Element == "SM") {
        p <- p_SM(p, t, chi, LR, alpha_rau)
        t <- t_WR(p, t, chi, LR)
      }
    }, error = function(e) {
      cat("Error in processing element:", Element, "at row", i, "in Kammer", kammer, "\n")
      cat("→", conditionMessage(e), "\n")
    })
    
    p2 <- p  # Save updated p after processing
    
    
    detailed_intermediate_values <- rbind(detailed_intermediate_values, data.frame(
      Kammer = kammer,
      Row = i,
      Element = Element,
      P_out = P_out,
      chi = chi,
      LR = LR,
      alpha_rau = alpha_rau,
      alpha_wi = alpha_wi,
      F1 = F1,
      F2 = F2,
      ds = ds,
      p1 = p1,
      p2 = p2,
      t = t,
      tau = tau,
      stringsAsFactors = FALSE
    ))
  }
  
  intermediate_results <- rbind(intermediate_results, data.frame(
    Kammer = kammer,
    p = p,
    t = t,
    p1=p1,
    p2=p2,
    F1 = F1,
    F2 = F2,
    tau = tau,
    ds = ds,
    stringsAsFactors = FALSE
  ))
}

timestamp <- format(Sys.Date(), "%Y%m%d")

summary_file <- paste0("ergebnisse/intermediate_results_", timestamp, ".csv")
write.csv(intermediate_results, file = summary_file, row.names = FALSE)
cat("Final summary saved to", summary_file, "\n")

detailed_file <- paste0("ergebnisse/detailed_intermediate_values_", timestamp, ".csv")
write.csv(detailed_intermediate_values, file = detailed_file, row.names = FALSE)
cat("Detailed intermediate values saved to", detailed_file, "\n")



#---- Abschnitt 1: Version _2 ----#

#---- Abschnitt 1 mit alle parameterwerte speicherung_bessere Version ----
if (!dir.exists("ergebnisse")) {
  dir.create("ergebnisse")
}

intermediate_results <- data.frame()
detailed_intermediate_values <- data.frame()
final_ratio_output <- data.frame()

kammer <-102
unique_kammer_values<-102
for (kammer in unique_kammer_values) {
  
  match_index <- which(Kammern$Kammer == kammer)
  kammer_Stollen <- Stollen[Stollen[[1]] %in% kammer, ]
  p <- kammer_Stollen$Po[1]  
  T <- kammer_Stollen$Tipo[1]  
  F1 <- NA
  F2 <- NA
  tau <- NA
  ds <- NA
  
  kammer_values <- data.frame()  
  p_prev <- NA
  T_prev <- NA
  # i<-1
  for (i in 1:nrow(kammer_Stollen)) {
    cat("Processing row:", i, "in Kammer:", kammer, "\n")
    
    P_out <- kammer_Stollen$P_out[i]
    chi <- kammer_Stollen$chi[i]
    LR <- kammer_Stollen$Lr[i]
    if (is.na(LR)) {
      warning("LR is NA, setting to 0")
      LR <- 0
    }
    alpha_rau <- kammer_Stollen$alpha_rau[i]
    alpha_wi <- kammer_Stollen$alpha_wi[i]
    Element <- kammer_Stollen$Element[i]
    Lo <- kammer_Stollen$Lo[i]
    Lss <- kammer_Stollen$Lss[i]
    F1 <- kammer_Stollen$F1[i]
    F2 <- kammer_Stollen$F2[i]
    ds <- kammer_Stollen$ds[i]
    cat("Element being processed:", Element, "\n")
    cat("Input values → P_out:", P_out, " | p:", p, " | T:", T, " | chi:", chi,
        " | LR:", LR, " | alpha_rau:", alpha_rau, " | alpha_wi:", alpha_wi, "\n")
    
    p1 <- p
    t1 <- T
    
    tryCatch({
      if (p <= 0) {
        warning("p <= 0 before element processing, setting p = 1e-6")
        p <- 1e-6
      }
      if (is.nan(T)) {
        warning("T is NaN, setting T = 1")
        T <- 1
      }
      
      if (Element == "Verz_1") {
        print(str(list(p = p, T = T, F1 = F1, F2 = F2, chi = chi, LR = LR, alpha_rau = alpha_rau, Element = Element)))
        T <-  0.7*T_WR(p, T, chi, LR)
        tau<-alpha_rau*T
        p <- p_Verz_1(P_out, p,alpha_wi, Lo, Lss)
        
      } else if (Element == "Verz_2") {
        print(str(list(p = p, t = t, F1 = F1, F2 = F2, chi = chi, LR = LR, alpha_rau = alpha_rau, Element = Element)))
        T <- 0.7*T_WR(p, T, chi, LR)
        tau<-alpha_rau*T
        p <- p_Verz_2(P_out, p, Lo, Lss)
        
      } else if (Element == "Verz_3") {
        print(str(list(p = p, t = t, F1 = F1, F2 = F2, chi = chi, LR = LR, alpha_rau = alpha_rau, Element = Element)))
        T <-  0.7*T_WR(p, T, chi, LR)
        tau<-alpha_rau*T
        p <- p_Verz_3(P_out, p, alpha_wi, Lo, Lss)
        
      } else if (Element == "Stollen") {
        print(str(list(p = p, t = t, F1 = F1, F2 = F2, chi = chi, LR = LR, alpha_rau = alpha_rau, Element = Element)))
        T <- T_WR(p,T,chi, LR)
        tau<-alpha_rau*T
        p <- p_WR(p, tau, chi, LR)
        
      } else if (Element == "PEW") {
        print(str(list(p = p, t = t, F1 = F1, F2 = F2, chi = chi, LR = LR, alpha_rau = alpha_rau, Element = Element)))
        T <- T_PEW(T, F1, F2)
        tau<-alpha_rau*T
        p <- p_PEW(p, F1, F2)
        
      } else if (Element == "PVE") {
        print(str(list(p = p, t = t, F1 = F1, F2 = F2, chi = chi, LR = LR, alpha_rau = alpha_rau, Element = Element)))
        T <- T
        tau<-alpha_rau*T
        p <- p_PVE(p, F1, F2)
        
      } else if (Element == "BL") {
        print(str(list(p = p, t = t, F1 = F1, F2 = F2, chi = chi, LR = LR, alpha_rau = alpha_rau, Element = Element)))
        T <- T
        tau<-alpha_rau*T
        p <- p_BL(p, F1, F2)
        
      } else if (Element == "SM") {
        print(str(list(p = p, t = t, F1 = F1, F2 = F2, chi = chi, LR = LR, alpha_rau = alpha_rau, Element = Element)))
        T <- T
        tau<-alpha_rau*T
        p <- p_WR(p, tau, chi, LR)
      }
      
      if (p <= 0) {
        warning("p <= 0 after processing, setting p = 1e-6")
        p <- 1e-6
      }
      
    }, error = function(e) {
      cat("Error in processing element:", Element, "at row", i, "in Kammer", kammer, "\n")
      cat("→", conditionMessage(e), "\n")
    })
    
    
    p2 <- p
    t2 <- T
    

    if (i == 1) {
      p_ratio <- 1
      t_ratio <- 1
    } else {
      p_ratio <- round(p2 / p_prev, 3)
      t_ratio <- round(t2 / t_prev, 3)
    }
    
    p_prev <- p2
    t_prev <- t2
    

    kammer_values <- rbind(kammer_values, data.frame(
      Kammer = kammer,
      Element = Element,
      p2 = round(p2, 6),
      t = round(t2, 4),
      p2_p1 = p_ratio,
      t2_t1 = t_ratio,
      stringsAsFactors = FALSE
    ))
    

    detailed_intermediate_values <- rbind(detailed_intermediate_values, data.frame(
      Kammer = kammer,
      Row = i,
      Element = Element,
      P_out = P_out,
      chi = chi,
      LR = LR,
      alpha_rau = alpha_rau,
      alpha_wi = alpha_wi,
      F1 = F1,
      F2 = F2,
      ds = ds,
      p1 = p1,
      p2 = p2,
      t = t2,
      tau = alpha_rau * t2,
      stringsAsFactors = FALSE
    ))
  }
  

  p_prod <- round(prod(kammer_values$p2_p1, na.rm = TRUE), 3)
  t_prod <- round(prod(kammer_values$t2_t1, na.rm = TRUE), 3)
  last_p <- kammer_values$p2[nrow(kammer_values)]
  last_t <- kammer_values$t[nrow(kammer_values)]
  

  kammer_values <- rbind(kammer_values, data.frame(
    Kammer = kammer,
    Element = "Schlüsselwerte",
    p2 = last_p,
    t = last_t,
    p2_p1 = p_prod,
    t2_t1 = t_prod,
    stringsAsFactors = FALSE
  ))
  

  final_ratio_output <- rbind(final_ratio_output, kammer_values)
}

timestamp <- format(Sys.Date(), "%Y%m%d")
write.csv(final_ratio_output, paste0("ergebnisse/final_ratio_output_", timestamp, ".csv"), row.names = FALSE)
cat("Correct ratio output saved with accurate Schlüsselwerte.\n")


#---- Abschnitt 2: Exposition ----
final_results <- data.frame()

for (kammer in unique(intermediate_results$Kammer)) {
  

  kammer_data <- intermediate_results[intermediate_results$Kammer == kammer, ]

  exposition_kammer <- unique(Exposition[Exposition$Kammer == kammer, ])
  if (nrow(exposition_kammer) == 0) {
    cat("No exposition data found for Kammer:", kammer, "\n")
    next
  }

  p <- kammer_data$p
  t <- kammer_data$t
  #lambda hier berechnen
  F1 <- kammer_data$F1
  F2 <- kammer_data$F2
  

  for (j in 1:nrow(exposition_kammer)) {
    Stollenmund <- exposition_kammer$Stollenmund[j]
    Objekt_Nr <- exposition_kammer$Objekt_Nr[j]
    Mund_X <- exposition_kammer$Mund_X[j]
    Mund_Y <- exposition_kammer$Mund_Y[j]
    Objekt_X <- exposition_kammer$Objekt_X[j]
    Objekt_Y <- exposition_kammer$Objekt_Y[j]
    Stollen_X <- exposition_kammer$Stollen_X[j]
    Stollen_Y <- exposition_kammer$Stollen_Y[j]

    d <- sqrt((Objekt_X - Mund_X)^2 + (Objekt_Y - Mund_Y)^2)
    numerator <- (Objekt_X - Mund_X) * (Mund_X - Stollen_X) + (Objekt_Y - Mund_Y) * (Mund_Y - Stollen_Y)
    denominator <- sqrt((Mund_X - Stollen_X)^2 + (Mund_Y - Stollen_Y)^2) * sqrt((Objekt_X - Mund_X)^2 + (Objekt_Y - Mund_Y)^2)
    omega <- acos(numerator / denominator)
    omega_degrees <- omega * (180 / pi)
    alpha <- asin((0.43 / 0.57) * sin(omega))
    alpha_degrees <- alpha * (180 / pi)
    gamma <- 180 - alpha_degrees - omega_degrees
    r <- sqrt(d^2 / (0.43^2 + 0.57^2 - 2 * 0.43 * 0.57 * cos(gamma * pi / 180)))
    p_final <- p / ((r / (0.7 * ds)) / 0.9)
    t_final <- p^(5/3) * t * (0.36 * ds / r)^2

    final_results <- rbind(final_results, data.frame(
      Kammer = kammer,
      Stollenmund = Stollenmund,
      Objekt_Nr = Objekt_Nr,
      Stollen_X = Stollen_X,
      Stollen_Y = Stollen_Y,
      Mund_X = Mund_X,
      Mund_Y = Mund_Y,
      Objekt_X = Objekt_X,
      Objekt_Y = Objekt_Y,
      Po = p,
      Tipo = t,
      p_final = p_final,
      t_final = t_final,
      stringsAsFactors = FALSE
    ))
  }
  
  cat("Processed Kammer:", kammer, "p:", p, "t:", t, "p_final:", p_final, "t_final:", t_final, "\n")
}

final_results <- unique(final_results)

timestamp <- format(Sys.Date(), "%Y%m%d")
final_file <- paste0("ergebnisse/final_results_", timestamp, ".xlsx")

if (file.exists(final_file)) {
  file.remove(final_file)
  cat("Old file removed.\n")
}

write.xlsx(final_results, file = final_file, rowNames= FALSE)
cat("Final results saved to", final_file, "\n")



 
