# ---------------------------------------------------------------------------------
# Instalacao de pacotes para paralelizar em caso de rede grande ----------------
# vide https://github.com/cwatson/brainGraph/issues/15
# ---------------------------------------------------------------------------------
# instalar apenas 1 vez
# caso ja tenha instalaso, inserir o  " # " antes do comando
install.packages('snow')
install.packages('doSNOW')
##################################################################################
# ---------------------------------------------------------------------------------
# Instalacao de pacotes para trabalhar com redes ----------------------------------
# ---------------------------------------------------------------------------------
# Instalar o pacote 'igraph', sua bib e os pacotes internos a 'igraph' 
# para trabalhar com redes

#install.packages("igraph")
library(igraph)
  
# Instalar o pacote 'brainGraph', sua bib e os pacotes internos a 'brainGraph' 
# para trabalhar com redes
# Cuidado! Pode haver conflito com o pacote "sand"

#install.packages("brainGraph")
library(brainGraph)

# Instalar o pacote 'sand', sua bib e os pacotes internos a 'sand' 
# para trabalhar com redes
# Cuidado! Pode haver conflito com o pacote "brainGraph"

#install.packages("sand")
#library(sand)
#install_sand_packages()

###################################################################
# ---------------------------------------------------------------------------------
# Executar esta parte para paralelizar
# ---------------------------------------------------------------------------------

OS <- .Platform$OS.type
if (OS == 'windows') {
  library(snow)
  library(doSNOW)
  num.cores <- as.numeric(Sys.getenv('NUMBER_OF_PROCESSORS'))
  cl <- makeCluster(num.cores, type='SOCK')
  clusterExport(cl, 'sim.rand.graph.par')   # Or whatever function you will use
  registerDoSNOW(cl)
} else {
  library(doMC)
  num.cores <- detectCores()
  registerDoMC(num.cores)
}  




# ---------------------------------------------------------------------------------
# Comandos para abrir arquivos de redes no formato Pajek --------------------------
# ---------------------------------------------------------------------------------

# Selecionar a rede de interesse no formato .NET (Pajek)
g <- read.graph(file.choose(), format = "pajek")

##############################################################################

# Calcular a eficiencia local da rede de interesse
# para grandes redes deve-se indicar como TRUE o comando use.parallel
##############################################################################

# Calcular a eficiência global da rede de interesse
# efficiency(g, type = c("local", "nodal", "global"), weights = NULL, use.parallel = TRUE, A = NULL)
#efficiency(g, type = "global", weights = NULL, use.parallel = FALSE, A = NULL)
# ou
#efficiency(g, type = "nodal", weights = NULL, use.parallel = FALSE, A = NULL)

enodal <- efficiency(g, type = "nodal", weights = NULL, use.parallel = TRUE, A = NULL)


mediaglobal<-mean(enodal)
mediaglobal

##############################################

# Calcular a eficiência local da rede de interesse
#efficiency(g, type = "local", weights = NULL, use.parallel = FALSE, A = NULL)

elocal <- efficiency(g, type = "local", weights = NULL, use.parallel = TRUE, A = NULL)
#elocal <- efficiency(g, type = "local", weights = NULL, use.parallel = FALSE, A = NULL)
medialocal<-mean(elocal)
medialocal

medialocal

 