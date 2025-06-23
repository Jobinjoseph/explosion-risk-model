# TW_functions.R - Function für explosion Berechnungen

#---- clean workspace ----
tw_clean <- function() {
  rm(list=ls())
}

#---- read Excel ----
read_explosion_data <- function(file_path, sheet_name) {
  library(readxl)
  data <- read_excel(file_path, sheet = sheet_name)
  return(data)
}
#---- berechnung dk und do # Ref: TLM Luftstoss Zugangsstollen A3-8 ---- 

dk_do <- function(data) {
  data <- data %>%
    mutate(
      dk = sqrt((4 * Fk) / pi),
      do = sqrt((4 * Fo) / pi)
    )
  return(data)
}

#---- berechnung LR  # Ref: TLM Luftstoss Zugangsstollen A3-7, Figur A3-3----
LR <- function(data) {
  data <- data %>%
    mutate(
      Lr = Ls - 5 * ds,
      Lr = case_when(
        Lr <= 0 ~ 0,
        TRUE ~ Lr
      )
    )
  return(data)
}

#---- Berechnung Po (atü) und Tipo (ms) # Ref: TLM Luftstoss Zugangsstollen A3-5,6, Figur A3-1,2----
Po_Tipo <- function(data) {
  data <- data %>%
    mutate(
      Po = case_when(
        Klotz == "nein" ~ ((400 * (Q / Vk)^(2/3) * (Lk / do)^(1/3) + 1) * 1.01325) / 10,# atü -> Mpa
        Klotz == "ja" & (0.6 < dB / do & dB / do < 0.7) ~ ((180 * (Q / Vk)^(2/3) * (Lk / do)^(1/3) + 1) * 1.01325) / 10,
        Klotz == "ja" ~ ((240 * (Q / Vk)^(2/3) * (Lk / dB)^(1/3) * (dB / do)^(1/3) + 1) * 1.01325) / 10,
        TRUE ~ NA_real_
      ),
      Tipo = case_when(
        Klotz == "nein" ~ 20 * Lk^(2/3) * do^(1/3) * (dk / do)^2,
        Klotz == "ja" & (0.6 < dB / do & dB / do < 0.7) ~ 8 * (Vk / Q)^(0.6),
        Klotz == "ja" ~ 1.7 * (Vk / Q)^(0.6) * sqrt(S) * 1, 
        TRUE ~ NA_real_
      )
    )
  return(data)
}


# ---- Berechnung Alpha (Wandrauhigkeitsbeiwert) für Stollenabschnitte # Ref: TLM Luftstoss Zugangsstollen A3-8, Figur A3-4----
alpha_rau_Stollen <- function(data) {

  alpha_rau <- rep(NA, nrow(data))

  alpha_rau[data$ds >= 2.5 & data$ds <= 3.5 & data$W == "Beton"] <- 1
  alpha_rau[data$ds >= 2.5 & data$ds <= 3.5 & data$W == "Beton/Backstein"] <- 1
  alpha_rau[data$ds >= 2.5 & data$ds <= 3.5 & data$W == "Gunit"] <- 4
  alpha_rau[data$ds >= 2.5 & data$ds <= 3.5 & data$W == "rohen Fels"] <- 6
  alpha_rau[data$ds >= 2.5 & data$ds <= 3.5] <- 1  

  alpha_rau[is.na(alpha_rau)] <- 1 * (2.8 / data$ds[is.na(alpha_rau)])^(4/3)

  data$alpha_rau <- alpha_rau
  
  return(data)
}





#---- berechnung Chi (Abstandkoeffizient in m) # Ref: TLM Luftstoss Zugangsstollen A3-8, Figur A3-4----
chi <- function(data) {
  data <- data %>%
    mutate(
      chi = alpha_rau*Lr
      
    )
  return(data)
}

#---- berechnung tau (Zeitdauerkoeffizient in ms) # Ref: TLM Luftstoss Zugangsstollen A3-8, Figur A3-4----
tau <- function(data) {
  data <- data %>%
    mutate(
      tau = alpha_rau*Tipo
    )
  return(data)
}


#---- Berechnung der Wandreibung #  Ref: TLM Luftstoss Zugangsstollen A8-11, Figur A3-4 ----
#---- Berechnung Wandreibung P---- 
# p <- 9.390478 
# t<- 1256.152
# tau<-0.001
p_WR <- function(p, tau, chi, LR) {
  cat("Inside p_WR → p:", p,  " | tau:", tau, " | chi:", chi, " | LR:", LR, "\n")
  if (LR == 0) {
    return(p)
  } else {
    t <- log((1000 / tau) + 1)
    m <- 0.01818 * t^2 + 0.16387 * t - 0.08809
    Z <- if (LR <= 1) 1 else (log(chi) + 6.35 * m) / (1 + m)
    cat("→ Z:", Z, "\n")
    
    P2_P1_ratio <- dplyr::case_when(
      LR <= 0 ~ 1,
      Z < 1 ~ 1,
      Z <= 3.254 ~ 0.097987 * Z^4 - 0.7498 * Z^3 + 1.8570 * Z^2 - 1.9977 * Z + 1.7929,
      Z > 3.254 ~ 0.3254 / Z,
      TRUE ~ NA_real_
    )
    
    cat("→ P2/P1 ratio:", P2_P1_ratio, "\n")
    p_new <- p * P2_P1_ratio
    cat("p after scaling:", p_new, "\n")
    return(p_new)
  }
}

T_WR <- function(p, T, chi, LR) {
  cat("Inside t_WR → p:", p, " | t:", T, " | chi:", chi, " | LR:", LR, "\n")
  if (LR < 0){
    return(T)
  } else {
    Z <- sqrt(p) / 40
    a <- -0.6415 * Z^3 + 2.6238 * Z^2 + 1.1953 * Z - 3.1192
    cat("→ Z:", Z, " | a:", a, "\n")
    
    T2_T1_ratio <- dplyr::case_when(
      LR <= 0 ~ 1,
      TRUE ~ 1 + (1 / 10000) * (a / 1000) * (chi^2) * (a + 8) * chi
    )
    
    cat("→ T2/T1 ratio:", T2_T1_ratio, "\n")
    T_new <- T * T2_T1_ratio
    cat("t after scaling:", T_new, "\n")
    return(T_new)
  }
}

p_Verz_1 <- function(P_out, p,alpha_wi, Lo, Lss) {
  cat("Inside p_Verz_1 → P_out:", P_out, " | alpha_wi:", alpha_wi, "\n")
  if (Lo >= 2 * Lss) {
    scale_factor <- 0.9
  } else {
    scale_factor <- if (P_out == "P2") sqrt(1 - alpha_wi / 180) else sqrt(alpha_wi / 180)
  }
  cat("→ scale_factor:", scale_factor, "\n")
  p_new <- p * scale_factor
  cat("p after scaling:", p_new, "\n")
  return(p_new)
}

  p_Verz_2 <- function(P_out, p, Lo, Lss) {
  cat("Inside p_Verz_2 → P_out:", P_out, "\n")
  if (Lo >= 2 * Lss) {
    scale_factor <- 0.9
  } else {
    scale_factor <- if (P_out == "P2") 0.8 else 0.25
  }
  cat("→ scale_factor:", scale_factor, "\n")
  p_new <- p * scale_factor
  cat("p after scaling:", p_new, "\n")
  return(p_new)
  }
  
  p_Verz_3 <- function(P_out, p, alpha_wi, Lo, Lss) {
  cat("Inside p_Verz_3 → P_out:", P_out, " | alpha_wi:", alpha_wi, "\n")
  if (Lo >= 2 * Lss) {
    scale_factor <- 0.9
  } else {
    scale_factor <- if (P_out == "P2") 0.8 else 0.8 * sqrt(1 - alpha_wi / 180)
  }
  cat("→ scale_factor:", scale_factor, "\n")
  p_new <- p * scale_factor
  cat("p after scaling:", p_new, "\n")
  return(p_new)
  }
  
  p_Verz_4 <- function(P_out, p, alpha_wi, Lo, Lss) {
  cat("Inside p_Verz_4 → P_out:", P_out, " | alpha_wi:", alpha_wi, "\n")
  if (Lo >= 2 * Lss) {
    scale_factor <- 0.9
  } else {
    scale_factor <- if (P_out == "P2") {
      0.9 - 0.6 * (alpha_wi / 180)^2
    } else {
      0.2 + 0.6 * (alpha_wi / 180)^2
    }
  }
  cat("→ scale_factor:", scale_factor, "\n")
  p_new <- p * scale_factor
  cat("p after scaling:", p_new, "\n")
  return(p_new)
  }
  
  p_BL <- function(p, F1, F2) {
  cat("Inside p_BL → F1:", F1, " | F2:", F2, "\n")
  Z <- sqrt(F2 / F1)
  scale_factor <- (1 / 1000) * (-529.37 * Z^3 + 496.26 * Z^2 + 1038.20 * Z - 6.99)
  cat("→ Z:", Z, " | scale_factor:", scale_factor, "\n")
  p_new <- p * scale_factor
  cat("p after scaling:", p_new, "\n")
  return(p_new)
  }
  
  p_PEW <- function(p, F1, F2) {
  cat("Inside p_PEW → F1:", F1, " | F2:", F2, "\n")
  Z <- sqrt(F1 / F2)
  scale_factor <- (1 / 1000) * (2186 * Z^4 - 5285.35 * Z^3 + 4149.94 * Z^2 - 52.61 * Z - 1.23)
  cat("→ Z:", Z, " | scale_factor:", scale_factor, "\n")
  p_new <- p * scale_factor
  cat("p after scaling:", p_new, "\n")
  return(p_new)
  }
  
  T_PEW <- function(T, F1, F2) {
  cat("Inside T_PEW → F1:", F1, " | F2:", F2, "\n")
  scale_factor <- sqrt(F1 / F2)
  cat("→ T:", T, " | scale_factor:", scale_factor, "\n")
  T_new <- T * scale_factor
  cat("T after scaling:", T_new, "\n")
  return(T_new)
  }
  
  p_KEW <- function(p, F1, F2) {
  cat("Inside p_KEW → F1:", F1, " | F2:", F2, "\n")
  Z <- sqrt(F1 / F2)
  scale_factor <- (1 / 1000) * (1866.65 * Z^4 - 4127.46 * Z^3 + 2625.67 * Z^2 + 641.78 * Z - 4.43)
  cat("→ Z:", Z, " | scale_factor:", scale_factor, "\n")
  p_new <- p * scale_factor
  cat("p after scaling:", p_new, "\n")
  return(p_new)
  }
  
  p_PVE <- function(p, F1, F2) {
  cat("Inside p_PVE → p:", p, " | F1:", F1, " | F2:", F2, "\n")
  x <- F2 / F1
  cat("→ x = F2 / F1 =", x, "\n")
  
  scale_factor <- dplyr::case_when(
    x < 0.08 ~ 1.66,
    x >= 0.08 ~ (1 / 1000) * (345.81 * x^2 - 1087.69 * x + 1754.71)
  )
  cat("→ scale_factor:", scale_factor, "\n")
  
  p_new <- p * scale_factor
  cat("p after scaling:", p_new, "\n")
  return(p_new)
  }
  
  
  p_KVE <- function(p, F1, F2) {
  cat("Inside p_KVE → F1:", F1, " | F2:", F2, "\n")
  t_local <- sqrt(F2 / F1) - 0.08
  scale_factor <- (1 / 1000) * (-2049.44 * t_local^3 + 3565.65 * t_local^2 - 4059.30 * t_local + 3425.46)
  cat("→ t_local:", t_local, " | scale_factor:", scale_factor, "\n")
  p_new <- p * scale_factor
  cat("p after scaling:", p_new, "\n")
  return(p_new)
  }
  
  p_SM <- function(p,tau, chi, LR) {
  cat("Inside p_SM → calling p_WR\n")
  p_new <- p_WR(p, tau, chi, LR)
  if (p <= 0) {
    p <- 1e-6
  }
  cat("p after p_WR (inside p_SM):", p_new, "\n")
  return(p_new)
  }
  
  




#---- Berechnung T2 # Ref: TLM Luftstoss Zugangsstollen A8-12, Figur A3-5 ----
omega <- function(x, y, x1, y1, x2, y2) {
  numerator <- (x - x1) * (x1 - x2) + (y - y1) * (y1 - y2)
  denominator <- sqrt((x1 - x2)^2 + (y1 - y2)^2) * sqrt((x - x1)^2 + (y - y2)^2)
  omega <- acos(numerator / denominator)
  return(omega)
}

# ---- Calculation Functions ----






