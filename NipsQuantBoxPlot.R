nipsQuant <- read.table("/Users/m006703/NIPS/NIPSQuantResult.txt", header = TRUE)

# Boxplot of initial quant
boxplot(InitialQuant~RunName, data = nipsQuant, main = "Initial Quant", xlab = "Run", ylab = "ng/ul")

# Boxplot of final quant
boxplot(FinalQuant~RunName, data = nipsQuant, main = "Final Quant", xlab = "Run", ylab = "ng/ul")
