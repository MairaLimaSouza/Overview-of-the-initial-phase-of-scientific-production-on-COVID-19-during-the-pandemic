# bibliotecas para manipular base de dados
library(data.table)
library(tidyverse)
library(tibble)

# para manipular strings
library(stringr)
library(sjmisc)

# Bibliotecas para excel
library(readxl)
library(openxlsx)
library(dplyr)

# bibiloteca para manipular dados e executa comandos is.empty, is.character, is.number, etc
require(rapportools)

#install.packages("Xmisc")
require(Xmisc)

#install.packages("abjutils")
require(abjutils)

# chamada dos pacotes/bibliotecas para realizar analise combinatoria.
#install.packages("gtools")
require(gtools)

#para rede citacao
require(bibliometrix)

# abre base original no diretorio 
path = "C:\\Users\\Visitante\\Documents\\mybiblio\\WOS_PBL2_1006.xlsx"

#caso a planilha tenha mais de uma sheet, percorro todas.
# indico quais as colunas desejo que o dataset tenha

sheet <- loadWorkbook(path)
sheetNames <- sheets(sheet)
data<-as.data.frame(matrix(,ncol=0,nrow=0))
var_norteadora=c('Authors','Author.Full.Name','Document.Title','Publication.Name','Document.Type','Author.Keywords','Keywords.Plus.','Abstract',
'Author.Address','Reprint.Address','Cited.References','Cited.Reference.Count','Web.of.Science.Core.Collection.Times.Cited.Count',
'Total.Times.Cited.Count..Web.of.Science.Core.Collection..BIOSIS.Citation.Index..Chinese.Science.Citation.Database..Data.Citation.Index..Russian.Science.Citation.Index..SciELO.Citation.Index.',
'Usage.Count..Last.180.Days.','Usage.Count..Since.2013.','Publisher','International.Standard.Serial.Number..ISSN.',
'Electronic.International.Standard.Serial.Number..eISSN.','International.Standard.Book.Number..ISBN.',
'X29.Character.Source.Abbreviation','ISO.Source.Abbreviation','Publication.Date','Year.Published','Volume','Issue',
'Beginning.Page','Ending.Page','Article.Number','Digital.Object.Identifier..DOI.','Book.Digital.Object.Identifier..DOI.',
'Page.Count','Web.of.Science.Categories','Research.Areas','PubMed.ID','Open.Access.Indicator')
for(i in 1:length(sheetNames))
{
    tmp = openxlsx::read.xlsx(path, i, detectDates= TRUE)%>%data.frame()
    data<- rbind(data, tmp[var_norteadora])
}

# verifico o quantas linhas e colunas tem o dado
dim(data)

# percorro todo o dataset para inserir a coluna "COD_PROD"
for (j in 1:nrow(data)){
   data[j,"COD_PROD"]<-str_c("PROD_00",j)
}
# reorganizo as colunas para que a nova coluna seja a primeira
data<- data %>% select(COD_PROD, everything())

# obras 
Data_excluidos<-subset(data,(is.empty(Authors)|Authors=='[Anonymous]'))

#as obras excluidas foram salvas no dataset DataAutoresObra_excluidos
#08 obras foram excluidas
Data_excluidos

#novo dataset sem os autores com nome em branco é criado
data_v2<-subset(data,!is.empty(Authors))

#novo dataset sem os autores com nome anonimo
data_v2<-subset(data_v2,Authors!='[Anonymous]')

dim(data_v2)

# Criando nova versao do dado original.
# Novas linhas. 1749 linhas.
# Criando novo codigo para as novas linhas

data_v2 <- data_v2 %>% 
  select(Document.Title, Year.Published,Abstract,Author.Keywords,Publication.Name,
         Reprint.Address,Digital.Object.Identifier..DOI.,Cited.References,Author.Full.Name,Authors)

# percorro todo o dataset para inserir a coluna "NEWCOD_PROD"
# reorganizo as colunas para que a nova coluna seja a primeira
for (j in 1:nrow(data_v2)){
   data_v2[j,"NEWCOD_PROD"]<-str_c("PROD_00",j)
}
data_v2<- data_v2 %>% select(NEWCOD_PROD, everything())

#removo acentos e pontuacoes diferentes do DOI
data_v2$Digital.Object.Identifier..DOI.<-data_v2$Digital.Object.Identifier..DOI.%>%abjutils::rm_accent()

# faço uma selecao para idenitificar se ha mais de uma ocorrencia de um mesmo DOI
# exibido o resultado
df<- data_v2%>%group_by(Digital.Object.Identifier..DOI.) %>%count(Digital.Object.Identifier..DOI.)%>% ungroup()%>%filter(n>1) # seleciono as obras duplicadas
df

lista<-df$Digital.Object.Identifier..DOI. # crio lista de DOI's duplicados
listadrop<-c("1")
lista<-lista[-as.integer(listadrop)] # excluo os casos em que DOI está vazio
lista

# armazeno as linhas que contêm o mesmo DOI
# mantem no dataset original apenas a primeira ocorrencia do DOI, deletando as demais.
    var_select_cod=list()
    data_doi_data_select<- data.frame( # dataset que coletara linhas de ocorrencia de um mesmo DOI 
                 DOI=character(),
                 r =list(),
                 COD=list(),
                 stringsAsFactors=FALSE) 
     for (j in seq(1,length(lista),1)){
        r <-which(lista[j]==data_v2$Digital.Object.Identifier..DOI.)
        if(length(r)>1) # so deleto se 
        {
        data_doi_data_select[j,"DOI"]<-lista[j]
        data_doi_data_select$r[j]<-list(which(lista[j]==data_v2$Digital.Object.Identifier..DOI.)) 
        data_doi_data_select$COD[j]<-list(data_v2[r,1])  
        linha<-r[2]
        data_v2 <- data_v2[-linha,]
        }
        else{
         if(length(r)==1) break()    
        } 
    }
data_doi_data_select

# apos a eliminacao da duplicidade, dataset ficou com 1740 artigos
dim(data_v2)

names(data_v2)

################################################################################
### função para rede coautoria
# funcao que abrevia nome de autor. Requer ajustes (v2 funcao criada)
###
################################################################################

# Uso regex para formatar nome de autores. A regex extrai as iniciais maiúsculas dos nomes a partir da vírgula.
# cria um de-para para nomes que estão no seguinte formato: SOBRENOME, NOME. 
# De: DE SOUZA, MAIRA LIMA -> Para: DE SOUZA, M.L.
# OBS: lista de autores tem nomes como: MAIRA LIMA DE SOUZA. Nestes casos, a regex apenas extrai as iniciais maiúsculas. 
# versão regex que também funciona <dados2[j,coluna_para]<-regmatches(i, gregexpr("^.+?,\\s[A-Z]+|\\s[A-Z]",i,perl=TRUE))>

abreviar_nome_autor <- function(autor) {
 
    lista_preposicoes = c('DA', 'DO', 'DAS', 'DOS', 'DE')
    lista_abreviacao = list()
    i=autor
    partes_nome<- str_split(i,"",simplify = TRUE) #separa o nome char por char ex: "RICHARDE" -> "R" "I" "C" "H" "A" "R" "D" "E"
    
    #print(paste0("autor: ",i))
    
    if(',' %in% partes_nome){
        
        #abreviacao<-str_trim(unlist(str_extract_all(i, "^.+?,\\s[A-Z]+|\\s[A-Z]"))) #regex não funciona quando não há espaço entre letras maiusculas
        
        #nova regex. Considera todas as maiusculas (juntas ou separadas) que aparecem no nome
        #nome=Lima,Maira Souza -> Lima,M S \ Lima, Maira Lima ->Lima, M L| Lima,ML-> Lima,ML
        
        abreviacao<-str_trim(unlist(str_extract_all(i, "^.+?,|[A-Z]+|\\s[A-Z]+"))) 
        if(nchar(abreviacao[2])>1){ # se nome tiver mais de uma letra eg: ML
           partes_nome2<- str_split(abreviacao[2],"",simplify = TRUE)
           abreviacao[2]<-str_c(paste0(partes_nome2,collapse="."),".")
        }       
        else{
           abreviacao[2]<-str_c(paste0(abreviacao[2],collapse="."),".")
        } 
        abreviacao[1]<-ifelse(str_contains(abreviacao[1], "-"),str_trim(str_replace(abreviacao[1], "-", "")), str_trim(abreviacao[1])) 
        
  #      lista_abreviacao<-str_c(paste0(abreviacao,collapse="."),".") 
        lista_abreviacao<-str_c(abreviacao[1],abreviacao[2]) 
    
    } #fecha o comando if(str_find(',', partes_nome)
        
    else{
                
                #caso o nome do autor não esteja marcado com vírgula, pegar o último sobrenome e abreviar o resto
                #lista_abreviacao.append(autor)
                      
                lista_partes_nome_autor = str_split(i,' ', simplify =TRUE) # gerando uma lista de partes do nome do autor
                
                
                # pegando o último nome para usar como parte não abreviada
                tam_nome<-length(lista_partes_nome_autor)
                
                if (tam_nome>1){ # nome deve ter tam > 1 para ser abreviado
                    
                    abreviacao=str_c(lista_partes_nome_autor[tam_nome-1]) #evitando de abreviar preposicoes 
                    if (abreviacao %in% lista_preposicoes){
                        abreviacao<-str_trim(str_c(abreviacao,lista_partes_nome_autor[tam_nome],","))
                        parte2<-lista_partes_nome_autor[1:tam_nome-2]
                        
                    }
                
                    else{
                        abreviacao=str_c(lista_partes_nome_autor[tam_nome])
                        abreviacao<-str_c(abreviacao,"," )
                        parte2<-lista_partes_nome_autor[1:tam_nome-1]
                        
                    }
                   
                
                    #coloco nome na lista parte2 e formato SOBRENOME,N1.N2.
                
                    for (k in seq(1,length(parte2),1)){
                
                        parte = parte2[k]
                        parte_aux=str_split(parte,"", simplify =TRUE)
                        parte=str_c(parte_aux[1],'.')#abreviando 
                        abreviacao=str_c(abreviacao,parte)
                                
                    } # fecha a repeticao para parte2
                 
                 lista_abreviacao<-cbind(lista_abreviacao,abreviacao) 
                 
                } # fecha condicao que define nome deve ter tam >1 apara ser abreviado
                else# se o nome=tam1. condicao que recupera autores sem sobrenome, ex: Umesh; Kundu, Debanjan; Selvaraj, Chandrabose; Singh, Sanjeev Kumar; Dubey, Vikash Kumar. 
                    # umesh é solto.n tem sobrenome p tratar. Apos o tratamento o UMESH permanece. sem esse  else, o umesh some.
                    
                {
                 abreviacao=lista_partes_nome_autor
                 lista_abreviacao<-cbind(lista_abreviacao,abreviacao)  
                }
                
   } #fecha else da condicao em que autor não esteja marcado com vírgula
                                 
   
    return(lista_abreviacao)
}

###########################################################################

# Funçoes para tratar titulo e gerar rede semantica
# Funcao 1: ajustar_nomes = remove espaços em barnco, acentos e caracteres especiais
# Funcao 2: ajustar_nomes2 = trata as palavras ligadas a covid19
# Funcao 3: int_to_words,chamada dentro da funcao 2 = transforma digito numerico em texto. digito 1 ->one
# Funcao 4: .simpleCap = deixa a primeira letra do titulo maiuscula e o restante em minusculo.

###########################################################################

#http://www.botanicaamazonica.wiki.br/labotam/doku.php?id=bot89:precurso:1textfun:inicio
#https://gomesfellipe.github.io/post/2017-12-17-string/string/
#install.packages("abjutils")

require(abjutils)
ajustar_nomes=function(titulo){
  novo<-titulo%>%
  stringr::str_trim() %>%                        #Remove espaços em branco sobrando
  stringr::str_to_lower() %>%                    #Converte todas as strings para minusculo
  abjutils::rm_accent() %>%                                #Remove os acentos com a funcao criada acima
  stringr::str_replace_all("[,*:&/' '!.();?]", " ") %>%     #Substitui os caracteres especiais por vazio
  stringr::str_replace_all("_+", " ")      #Substitui os caracteres especiais por "_"   
  #  gsub("-","M",txt)      #Substitui o caracter especiais por "_"
 return(novo)  
}



ajustar_nomes_2=function(x){
    
    partes_nome<- str_split(toupper(x)," ")[[1]] # separo o titulo em palavras
    partes_nome
    vc_covid<-c("2019-NOVEL","2019-NOVEL-CORONAVIRUS","COVID-19","COVID19","2019-NCOV","SARS-NCOV-2","SARS-COV-2","SARSCoV-2","CORONAVIRUS-2") 

    for (k in seq(1,length(partes_nome),1)){
        parte = partes_nome [k]
        
        if(nchar(parte)==1 && str_detect(partes_nome [k],"[a-z|A-Z]") ){ # retira letras soltas apos eliminacao apostrofo
           partes_nome [k]=""
           next  
        }
        
        if(parte=="CORONAVIRUS" && k<length(partes_nome)) {
            if ( partes_nome [k+1]=="2"){
                # se tiver coronavirus 2
            partes_nome [k]="COVIDONENINE"
            partes_nome [k+1]=""
            next
            }
        }
                  
        if(parte=="NON-COVID-19") { # se tiver a frase novel coronavirus
            partes_nome [k]="NON COVIDONENINE"
            next
        }
        
        if(parte=="POST-COVID-19") { # se tiver a frase novel coronavirus
            partes_nome [k]="POST COVIDONENINE"
            next
        }
        
        if(parte=="2019-NOVEL" && k<length(partes_nome) && partes_nome [k+1]=="CORONAVIRUS") { # se tiver a frase novel coronavirus
            partes_nome [k]="COVIDONENINE"
            partes_nome [k+1]=""
            next
        }
        
        if(parte=="2019" && k<length(partes_nome) && partes_nome [k+1]=="NOVEL" && partes_nome [k+2]=="CORONAVIRUS") { # se tiver a frase novel coronavirus
            partes_nome [k]="COVIDONENINE"
            partes_nome [k+1]=""
            partes_nome [k+2]=""
            next
        }
       if(parte=="SARS" && k<length(partes_nome) && partes_nome [k+1]=="CoV-2") { # se tiver a frase novel coronavirus
            partes_nome [k]="COVIDONENINE"
            partes_nome [k+1]=""
            next
        } 
         
        if(parte=="CORONAVIRUS" && k<length(partes_nome) && partes_nome [k+1]=="DISEASE" && partes_nome [k+2]=="2019") { # se tiver a frase novel coronavirus
            partes_nome [k]="COVIDONENINE"
            partes_nome [k+1]="DISEASE"
            partes_nome [k+2]=""
            next
        }
                        
        if(parte=="CORONA" && k<length(partes_nome) && partes_nome [k+1]=="DISEASE") { # se tiver a frase corona disease #v.20/06/2020
            partes_nome [k]="COVIDONENINE"
            partes_nome [k+1]=""
            next
        }
        if(parte=="CORONAVIRUS" && k<length(partes_nome) && partes_nome [k+1]=="DISEASE-19") { # se tiver a frase corona disease #v.20/06/2020
            partes_nome [k]="COVIDONENINE"
            partes_nome [k+1]="DISEASE"
            next
        }
        if(parte=="COVID" && k<length(partes_nome) && partes_nome [k+1]=="DISEASE") { # se tiver a frase covid disease #v.20/06/2020
            partes_nome [k]="COVIDONENINE"
            partes_nome [k+1]=""
            next
        }
        if(parte=="COVID" && k<length(partes_nome) && partes_nome [k+1]=="19") { # se tiver a frase covid disease #v.20/06/2020
            partes_nome [k]="COVIDONENINE"
            partes_nome [k+1]=""
            next
        }
        if(parte=="COVID-19-ASSOCIATED"|parte== "SARS-COV-2-ASSOCIATED") { # se tiver a frase covid disease #v.20/06/2020
            partes_nome [k]="COVIDONENINE ASSOCIATED"
            next
        }
        if(parte=="COVID-19-MOTIVATED") { # se tiver a frase covid disease #v.20/06/2020
            partes_nome [k]="COVIDONENINE MOTIVATED"
            next
        }
        if(parte=="COVID-19-RELATED") { # se tiver a frase covid disease #v.20/06/2020
            partes_nome [k]="COVIDONENINE RELATED"
            next
        } 
        if(parte=="COVID-19-RECOMMENDATIONS") { # se tiver a frase covid disease #v.20/06/2020
            partes_nome [k]="COVIDONENINE RECOMMENDATIONS"
            next
        }
              
        if(parte %in%vc_covid){ # se em cada parte encontrar alguma palavra que esteja no vc_covid
            partes_nome [k]="COVIDONENINE"
            next
        } 
        #parte do codigo que coloca numeros por extenso
        if(str_detect(partes_nome [k],"[0-9]"))
            {
            num<-gsub("[^0-9.]+", "", partes_nome [k])
            digito= int_to_words(as.numeric(gsub("[^0-9.]+", "", partes_nome [k])))
            if(str_detect(partes_nome [k],"-") && !str_detect(partes_nome [k],"[a-z|A-Z]")){
              consoante=gsub ("[^a-z|A-Z.]+","",partes_nome [k]) 
              partes_nome [k] = gsub("[[:space:]]", "", paste(toupper(digito),consoante,collapse=""))  
            }
            else{
                frase=str_replace(partes_nome[k],num,digito)
                partes_nome [k] = gsub("[[:space:]]", "", paste(toupper(frase),collapse=""))
            }
                 
            }

    } # fecha a repeticao 
    
       
    x<-paste(partes_nome, collapse=" ")
    
    # v.20/06/2020 eliminei o execesso de espaco entre as palavras
    novo_nome<-gsub("[[:space:]]","_",x) #todo espaco sera substituido por _
    novo_nome<-gsub('[\"]', '', novo_nome) # retira as aspas duplas
    
    titulofinal<-novo_nome%>%
      stringr::str_replace_all("_+", " ") #subsitituo todo o _ por " "
    
    return(titulofinal)
}    


.simpleCap <- function(x) { # coloco a primeira letra da frase em maisculo 
    s <- strsplit(x, " ")[[1]]
#   paste(toupper(substring(s, 1, 1)), substring(s, 2), sep = "", collapse = " ")
    paste(toupper(substr(x,1,1)),substring(x, 2), sep = "", collapse = " ") #https://stackoverflow.com/questions/15897236/extract-the-first-or-last-n-characters-of-a-string
}


#https://stackoverflow.com/questions/18332463/convert-written-number-to-number-in-r
# funcão que coloca o numero por extenso. numero 2020=twozerotwozero

int_to_words <- function(x) {
    
    numero<-list()
    digits <- str_split(as.character(x), "",simplify = TRUE)
    nDigits <- length(digits)
    
    words <- c('zero', 'one', 'two', 'three', 'four',
                'five', 'six', 'seven', 'eight', 'nine',
                 'ten')
    
    for (k in seq(1,nDigits,1)){
    parte = digits [k]
    index <- as.integer(digits) + 1
    numero<-words[index]   
    #numero<-cbind(numero,words[index])
    
       
    } # fecha a repeticao
    numero= paste(numero,collapse = "")  
    return(numero)
   
}

#################################################################################
# sequencia abaixo realiza:
# Como a coluna AUTHOR eh uma lista que contem os nomes dos autores, os nomes devem ser separados.
# com os nomes separados, executo a funcao que abrevia o nome dos autores
# cria o dataset que tera o vocabulario de controle do nome dos autores tratados
# cria o dataset que terá a lista de autores com nome tratado (abreviado) e será a base para a criacao da rede coautoria, 
# o dataset esta estruturado por coluna: OBRA | AUTOR1 | AUTOR 2 | AUTOR 3 ....
# e contera nome de autor abreviado e nomes de autores separados em coluna
#####################################################################################

head(data_v2)

# criar colunas de autores

j=0

DataDePara<-as.data.frame(matrix(,ncol=0,nrow=0)) # dataset que tera o vc do nome dos autores

DataComb<-data.frame(AUTHOR=list(), # dataset que terá a lista de autores com nome tratado (abreviado) e será a base para a criacao da rede co-autoria
                 YEAR=character(), 
                 NEWCOD_PROD=character(), 
                 stringsAsFactors=FALSE) 
DataPBL<-as.data.frame(matrix(,ncol=0,nrow=0)) # dataset que contera nome de autor abreviado e nomes de autores separados em coluna
DataExcluido<-as.data.frame(matrix(,ncol=0,nrow=0)) # dataset que conterá dados excluidos

linhaautor_drop<-list() # lista que conterá autores que deverão sair do dataset


for (j in 1:nrow(data_v2)){
      
    autoria<-c('') #lista que recebera coluna Authors
    autoria_aux<-c('')
    
    nomecompleto<-c('') #lista que recebera coluna Authors.Fullname
    nomecompleto_aux<-c('')
    
   #altero para pegar o nome completo    
    autoria_aux<-str_split(unlist(as.list(data_v2[j,11])), fixed(';')) #ponho autor em lista, depois separo por ;. resulta em uma lista (1 linha de nomes)
    autoria<-str_split(unlist(autoria_aux), fixed('""')) #separo autor por autor. resulta em uma lista com linhas=n autores
    
    nomecompleto_aux<-str_split(unlist(as.list(data_v2[j,10])), fixed(';')) #ponho autor em lista, depois separo por ;. resulta em uma lista (1 linha de nomes)
    nomecompleto<-str_split(unlist(nomecompleto_aux), fixed('""'))  
     
    count<-1
    totalautoresobra<-length(autoria)
    listadatacomb<-list()
    cod_prod<-data_v2[j,1]
    DataPBL[j,"NEWCOD_PROD"]<-cod_prod
    DataPBL[j,"TITULO"]<-data_v2[j,2]
    DataPBL[j,"ANO"]<-data_v2[j,3]
    DataPBL[j,"RESUMO"]<-data_v2[j,4]
    DataPBL[j,"PALAVRAS_CHAVES"]<-data_v2[j,5]
    DataPBL[j,"PUBLICACAO"]<-data_v2[j,6]
    DataPBL[j,"AFILIACAO"]<-data_v2[j,7]
    DataPBL[j,"DOI"]<-data_v2[j,8]
    DataPBL[j,"CR"]<-data_v2[j,9]
       
    for (i in autoria){ 
      
        if (is.na(i)| i==" "| i=="...") break # caso tenha autor sem nome, nulo, etc...
        
       # abrevio nome autor
        nomeabreviado<-toupper(abreviar_nome_autor(i))
        nomeabreviado<-gsub("[[:space:]]", "", nomeabreviado) #https://stackoverflow.com/questions/5992082/how-to-remove-all-whitespace-from-a-string
        
        coluna<-str_c("AUTOR_", count)
        coluna_para<-str_c("VC_AUTOR_",count)    
              
        DataPBL[j,coluna]<-nomeabreviado                  
    
        DataDePara[j,"NEWCOD_PROD"]<-cod_prod
        DataDePara[j,coluna]<-lstrip(toupper(nomecompleto[count]), char = " ")
        DataDePara[j,coluna_para]<-nomeabreviado
        
        
        listadatacomb=cbind(listadatacomb,nomeabreviado) #coleta os nomes abreviados p analise combinatoria
        
        count<-count+1
   }# fecha o loop dos autores contidos na lista autoria
   
    DataComb[j,"AUTHOR"]<-""
    DataComb$AUTHOR[j]<-list(listadatacomb) # alimento o dataset da analise combinatoria
    DataComb[j,"YEAR"]<-data_v2[j,3]
    DataComb[j,"NEWCOD_PROD"]<-cod_prod

}#fecha o loop (for) que percorre o df


} 

#https://stackoverflow.com/questions/11095992/generating-all-distinct-permutations-of-a-list-in-r?noredirect=1&lq=1
#https://davetang.org/muse/2013/09/09/combinations-and-permutations-in-r/ (melhor solucao, usei a fc de combinacao. A fc de permutacao gera repetiçoes)
####################################################
#### dataset para a rede de coautoria vira do dataset DataComb
#### aplico a funcao de combinacao 
####
#####################################################

# Dataset criado a partir do dataset DataComb
dados_combina_aut <- DataComb %>% 
  select(AUTHOR,YEAR,NEWCOD_PROD)%>% as.data.table()

# chamada dos pacotes/bibliotecas para realizar analise combinatoria.
#install.packages("gtools")
require(gtools)

matriz<-NULL
m<-matrix()
data_comb <- data.frame()
ano<-""

for (n in 1:nrow(dados_combina_aut)){
    listaautoria<-list()
    
    listaautoria<-toupper(str_split(unlist(dados_combina_aut[n,1]), fixed('""')))
    listaautoria<-unique(sort(listaautoria))
    totalautoresobra<-length(listaautoria)
        
    
    if(totalautoresobra>=2)
    {
        matriz=gtools::combinations(totalautoresobra,2,listaautoria,repeats=FALSE)
        v1<-matriz[,1]
        v2<-matriz[,2]
        ano<-as.character(dados_combina_aut[n,2])
        cod<-as.character(dados_combina_aut[n,3])
        v3<-rep(ano, times= length(v1))
        v4<-rep(cod, times= length(v1))
        m<-cbind(v1,v2,v3,v4)
        
    }
    else # há obras que tem apenas um autor.Nestes casos, o autor combina com ele mesmo.
        {
        matriz<-listaautoria
        v1<-matriz
        v2<-matriz
        ano<-as.character(dados_combina_aut[n,2])
        cod<-as.character(dados_combina_aut[n,3])
        v3<-rep(ano, times= length(v1))
        v4<-rep(cod, times= length(v1))
        m<-cbind(v1,v2,v3,v4)
        
        
    }
    data_comb<-rbind(data_comb,m)
    
    
}


names(data_comb)<-c("DE","PARA","ANO","NEWCOD_PROD")

#cabecalho do dataset de coautoria
head(data_comb)

#############################################
# codigo abaixo trata o titulo e gera o arquivo para rede semantica
# dataset DataTituloVC, coluna TITULO_A (titulo after tratamento)
############################################

j=0
#'NEWCOD_PROD''TITULO_B''TITULO_A'
DataTituloVC<-as.data.frame(matrix(,ncol=0,nrow=0)) # dataset que terá o Vc dos titulos

for (j in 1:nrow(DataPBL)){
 
  titulos<-list()
  titulos<-DataPBL[j,2]
  DataTituloVC[j,"NEWCOD_PROD"]<-DataPBL[j,1]  
  DataTituloVC[j,"TITULO_B"]<-lstrip(DataPBL[j,2],char = " ")
  
  titulo=ajustar_nomes(titulos) # chamo a funcao que elimina acentos e caracteres especiais e deixa as palavras em minúsculo 
  titulo=ajustar_nomes_2(titulo) # chamo a funcao que aplica as regras 3 e 4 (vide descricao script)
  titulo=paste(str_replace_all(titulo,"-",""), collapse="") # elimino espaço em branco adicional
  
  titulo=.simpleCap(tolower(titulo)) # chamo a funcao que aplica a regra 5
    
  titulos=cbind(titulos,titulo) #coloca o titulo em uma lista de titulos
    
  DataPBL[j,2]<-lstrip(titulo, char = " ") #elimino espaco esquerda
  DataTituloVC[j,"TITULO_A"]<-lstrip(titulo, char = " ") #elimino espaco esquerda
 
}

##########################################################################
### a criacao do dataset que coleta as citacoes, observa os seguintes fatores

# ha obras que nao têm a coluna "CR" preenchida
# estas obras ficam de fora do dataset que recolhe as citacoes

#Acoes:

# coleta das citacoes doi, autor e ano
# quando o doi nao existe um codigo doi eh criado
# o dataset criado "Data_CR" estruturado da seguinte forma: 
# doi(obra citada) -> citacao origem (COD_PROD_OR) | codigo criado para producao citada (COD_PROD_CIT)
# COD_PROD_CIT tem a estrutura codigoproducao +"_CIT_"+ contador de citacao da obra original(assim se a obra tem 19 citacoes, o contador identifica citacao 1 ate citacao 19)


#https://www.rdocumentation.org/packages/regexPipes/versions/0.0.1/topics/grep
start_time <- Sys.time()

novodoi=NULL
autor=NULL
ano=NULL

Data_CR<-data.frame( # dataset que terá a lista de autores com nome tratado (abreviado) e será a base para a criacao da rede co-autoria
                 COD_PROD_OR=character(),
                 COD_PROD_CIT=character(), 
                 AUTHOR=character(),
                 ANO=character(),
                 DOI=character(),          
                 stringsAsFactors=FALSE) 


listadoi=list()

i=1

for (j in seq(1,nrow(DataPBL),1)){

    
    autoria_aux=list()
    autoria=list()
    cat("'linha data \n\t", j, "\n")
    
    
    autoria_aux<-str_split(unlist(DataPBL[j,"CR"]), fixed(';')) #separo por ;. resulta em uma lista (1 linha de citacoes)
    autoria<-str_split(unlist(autoria_aux), fixed(',')) #separo autor por autor. resulta em uma lista com linhas=n autores    
    
    
    codigoproducao<-DataPBL[j,"NEWCOD_PROD"]
    
    if(is.empty(DataPBL[j,"CR"])){
     next
    }
    
    for (aut in autoria){
      
        Data_CR[i,"COD_PROD_OR"]<-codigoproducao
                 
        autor=aut[1] # recupera o nome do autor que esta na posicao 1
        autor=trimws(trimES(base::gsub("[[:punct:]]"," ",autor))) 
        ano=aut[2] #recupera o ano que esta na posicao 2
        
        if(length(k <- aut %>% regexPipes::grep("DOI"))){
            novodoi<-trim.leading(aut[k])
            
           if(length(novodoi)>1){
             novodoi<-trim.leading(novodoi[2])
           }
        novodoi<-novodoi %>% regexPipes::gsub("([DOI])", "_")
        novodoi<-novodoi%>%stringr::str_replace_all("_+", " ") #subsitituo todo o _ por " "
        novodoi<-gsub("\\[|\\]", "", novodoi) #remove-square-brackets-from-a-string-vector
        novodoi<-trim.leading(novodoi)

        }
        else{
          novodoi=str_c("sem doi","_CIT_",i)  
        }
       Data_CR[i,"COD_PROD_CIT"] <-str_c(codigoproducao,"_CIT_",i)
        
       Data_CR[i,"AUTHOR"]<-autor
        
       if (!is.empty(ano) && str_detect(ano,"[0-9]")&& !str_detect(ano,"[a-zA-Z]")){ #para os casos em que ano esta vazio        
           ano<-ano
       }
       else{
           ano<-""
       } 
        
       Data_CR[i,"ANO"]<-ano
       Data_CR[i,"DOI"]<-novodoi 
     
        i<-i+1 
  }#fecha for autoria
    
} #finaliza o for que percorre data

end_time <- Sys.time()
end_time - start_time


#############################
# necessario verificar a ocorrencia de DOI's duplicados em Data_CR

#Coleto em quais linhas cada DOI (unico) aparece no Data_CR.

#  verifico se há doi que se repete no Data_CR(citado por mais de 1 obra)
#  como todo doi está preenchido no dataset
#  a checagem é feita a partir da lista que coleta doi's únicos

count<-1
listadoi<-unique(Data_CR$DOI)
data_doi_select<- data.frame( # dataset que terá a lista de autores com nome tratado (abreviado) e será a base para a criacao da rede co-autoria
                 DOI=character(),
                 r =list(),                   
                 stringsAsFactors=FALSE) 
    
  
     for (j in seq(1,length(listadoi),1)){
         
        cat("\n tam listadoi ",j,"\n") 
        r <-which(listadoi[j]==Data_CR$DOI)
        cat("\n listadoi[j] ",listadoi[j],"\n")  
        cat("\n r ",r,"\n")
        data_doi_select[j,"DOI"]<-listadoi[j]
        data_doi_select$r[j]<-list(which(listadoi[j]==Data_CR$DOI))
         
    }

##########################
# gero dataset com a seguinte estrutura: Obra é citada por
# As obras neste dataset não estao com o COD_PROD
# A chave do dataset é o DOI

# dataset gerado a partir das linhas contidas no dataset data_doi_select 
start_time <- Sys.time()
count<-1
var_select_doi=NULL
var_select_cod_or=NULL
var_select_codcit=NULL
var_select_autor=NULL
var_select_ano=NULL

data_var_select<- data.frame( 
                 DOI=character(),
                 AUTHOR=character(),
                 ANO=character(),   
                 COD_OR=list(),
                 COD_CIR=list(),          
                 stringsAsFactors=FALSE) 
    
        
       for (j in 1:nrow(data_doi_select)){
         
        r <-unlist(data_doi_select[j,2])  
                      
             var_select_doi<-data_doi_select[j,1] 
             var_select_autor<-as.list(Data_CR[r,3])
             var_select_codcit<-as.list(Data_CR[r,2])
             var_select_cod_or<-as.list(Data_CR[r,1])
             var_select_ano<-as.list(Data_CR[r,4])
             
             data_var_select[count,"DOI"]<-var_select_doi
             data_var_select[count,"AUTHOR"]<-var_select_autor[1]
             data_var_select[count,"ANO"]<-var_select_ano[1]
           
             for (d in 1:length(r)){
                coluna<-str_c("COD_OR_", d)
                coluna_para<-str_c("COD_CIR_",d)    
                data_var_select[count,coluna]<-var_select_cod_or[d]                  
               # data_var_select[count,coluna_para]<-var_select_codcit[d]
                
            } #fecha o for

           count= count+1  
              
      
    } # fecha o for 

end_time <- Sys.time()
end_time - start_time

#importando o cod_prod dos dois existentes no dataset do vocabulario de controle de doi's

#####################################
# faco merge para inserir as obras no vocabulario de controle DataCitacaoVc
# data_var_select_v2 resulta do merge entre data_var_select e DataCitacaovc
# objetivo coletar o COD_PROD contidas das obras que aparecem em ambos os datasets (variavel comum DOI)
# depois checa a ocorrencia do doi recolhido pelo data_var_select é unico ou nao.

# data_var_select_v2 resulta do merge entre data_var_select e DataCitacaovc
# objetivo coletar o COD_PROD contidas das obras que aparecem em ambos os datasets (variavel comum DOI)
data_var_select_v2<-merge(data_var_select, DataCitacaovc, by.x="DOI", by.y="DOI", all.x=TRUE,sort = FALSE)

data_var_select_v2<-data_var_select_v2%>%select(COD_PROD,everything())

############################################
# Dataset que controla quais doi's estao em DataCitacaovc(obraoriginal) e aparecem no dataset de citacao 

# verifico se alguma obra contida DataCitacaovc existe em  data_var_select
# caso exista , recupero a linha em que o DOI aparece no data_var_select
listadoi<-unique(DataCitacaovc$DOI)

data_doi_citacao_comum<- data.frame( 
                 DOI=character(),
                 r =list(),                   
                 stringsAsFactors=FALSE) 
count=1
for (j in 1:length(listadoi)){
         
    cat("\n count tam listadoi ",j,"\n")
    
    if(!is.empty(listadoi[j])){
        
        r <-which(listadoi[j]==data_var_select_v2$DOI)
    
        if (length(r)==0){ 
            cat("\n nao ha elemento em comum\n")
        }
        else{
           if (length(r)>=1){
            cat("\n listadoi[j] ",listadoi[j],"esta em comum \n")  
            cat("\n r ",r,"\n")
            data_doi_citacao_comum[count,"DOI"]<-listadoi[j]
            data_doi_citacao_comum$r[count]<-list(which(listadoi[j]==data_var_select_v2$DOI))
            count=count+1
           } #fecha o if
        } #fecha o else     
        
    } #fecha if do empty
         
        
} #fecha o for



### Insiro as obras da citacao (data_var_select_v2) no DataCitacaovc
#### Acoes:
#### Ao inserir no datacitacaovc codigo eh criado para cada obra nova
#### importo para o data_var_select_v2 o codigo da obra gerado
#### novo codigo inicia a partir de 1750 (PROD_001749 eh o ultimo codigo de obra original no DataCitacaovc

# insere a obra do data_var_select_v2 n
tam=1750 #ultimo codigo é 1749 e acrescido do +1
listacontroladoi<-DataCitacaovc$DOI

for (j in seq(1,nrow(data_var_select_v2),1)){ #percorro data_var_select
    
    if(data_var_select_v2[j,"DOI"] %in% listacontroladoi){
        cat("\n DOI ",data_var_select_v2[j,"DOI"],"ja esta inserido\n")
       
    }
    else{
        
      DataCitacaovc[tam,"COD_PROD"]<-str_c("PROD_00",tam)
      DataCitacaovc[tam,"AUTHOR"]<-data_var_select_v2[j,"AUTHOR.x"]
      DataCitacaovc[tam,"ANO"]<-data_var_select_v2[j,"ANO.x"]
      DataCitacaovc[tam,"DOI"]<-data_var_select_v2[j,"DOI"]
      data_var_select_v2[j,"COD_PROD"]<-str_c("PROD_00",tam) # coloco o codigo novo gerado para o DataCitacaovc
       
      tam=tam+1
        
    }
      
  
}  
data_var_select_v2<- data_var_select_v2 %>% select(COD_PROD, everything())    

###########################################################
#Importando o codigo (COD_PROD) gerado para o dataset DataCitacaovc na obra correspondente contida no Data_CR.

#Realizacao do innerjoin by "DOI" (innerjoin mantem a ordem das linhas)
# Resultado aramazenado no dataset Data_CR_COD_ANTI
# Checagem do tamanho para verificar se a integridade se mantêm

#inner join mantem a ordem das linhas
Data_CR_COD_ANTI<-inner_join(Data_CR,data_var_select_v2,by="DOI")
Data_CR_COD_ANTI<-Data_CR_COD_ANTI%>%select(COD_PROD,DOI,COD_PROD_OR,COD_PROD_CIT,AUTHOR, ANO)

############################################################
# criacao da rede de citacao

lista_cod_or=unique(Data_CR_COD_ANTI$COD_PROD_OR)
data_cod_citacao_comum<- data.frame( 
                 COD_PROD_OR=character(),
                 r =list(),                   
                 stringsAsFactors=FALSE) 
count=1

for (j in seq(1,length(lista_cod_or),1)){
         
    cat("\n count tam listadoi ",j,"\n")
    
    if(!is.empty(lista_cod_or[j])){
        
        r <-which(lista_cod_or[j]==Data_CR_COD_ANTI$COD_PROD_OR)
    
        if (length(r)==0){ 
            cat("\n nao ha elemento em comum\n")
        }
        else{
           if (length(r)>=1){
            cat("\n listadoi[j] ",lista_cod_or[j],"esta em comum \n")  
            cat("\n r ",r,"\n")
            data_cod_citacao_comum[count,"COD_PROD_OR"]<-lista_cod_or[j]
            data_cod_citacao_comum$r[count]<-list(which(lista_cod_or[j]==Data_CR_COD_ANTI$COD_PROD_OR))
            count=count+1
           } #fecha o if
        } #fecha o else     
        
    } #fecha if do empty
         

}

###########################
# rede versao coluna

start_time <- Sys.time()
count<-1

var_select_doi=NULL
var_select_codprod=NULL
#var_select_cod_or=NULL
var_select_codcit=NULL
var_select_autor=NULL
var_select_ano=NULL

data_cocitacao<- data.frame( 
                 COD_PROD=character(),
                 stringsAsFactors=FALSE) 
    
        
       for (j in 1:nrow(data_cod_citacao_comum)){
         
        r <-unlist(data_cod_citacao_comum[j,2])  
                      
             var_select_codprod<-data_cod_citacao_comum[j,1]
             var_select_doi<-as.list(Data_CR_COD_ANTI[r,2])
             var_select_autor<-as.list(Data_CR_COD_ANTI[r,5])
             var_select_ano<-as.list(Data_CR_COD_ANTI[r,6])
             var_select_cod_or<-as.list(Data_CR_COD_ANTI[r,1])
             var_select_codcit<-as.list(Data_CR_COD_ANTI[r,4])
           
             data_cocitacao[count,"COD_PROD"]<-var_select_codprod  
             #data_cocitacao[count,"DOI"]<-var_select_doi[1]
                          
             for (d in 1:length(r)){
                coluna<-str_c("TAG_", d)
                #coluna_ano<-str_c("ANO_", d) 
                coluna_para<-str_c("COD_CIT_",d)
                #coluna_doi<-str_c("DOI_CIT_",d) 
                data_cocitacao[count,coluna_para]<-var_select_cod_or[d]
                #data_cocitacao[count,coluna]<-var_select_autor[d]
                data_cocitacao[count,coluna]<-str_c(var_select_autor[d],var_select_ano[d])
                #data_cocitacao[count,coluna_ano]<-var_select_ano[d] 
                #data_cocitacao[count,coluna_doi]<-var_select_doi[d] 
                                
            } #fecha o for

           count= count+1  
              
      
    } # fecha o for 

end_time <- Sys.time()
end_time - start_time

###########################
# rede versao linha

start_time <- Sys.time()
count<-1

var_select_doi=NULL
var_select_codprod=NULL
#var_select_cod_or=NULL
var_select_codcit=NULL
var_select_autor=NULL
var_select_ano=NULL

data_coc<- data.frame( 
                 COD_PROD=character(),
                 stringsAsFactors=FALSE) 
    
        
       for (j in 1:nrow(data_cod_citacao_comum)){
         
        r <-unlist(data_cod_citacao_comum[j,2])  
                      
             var_select_codprod<-data_cod_citacao_comum[j,1]
             var_select_cod_or<-as.list(Data_CR_COD_ANTI[r,1])
             var_select_codcit<-as.list(Data_CR_COD_ANTI[r,4])
             
             var_select_doi<-as.list(Data_CR_COD_ANTI[r,2])
             var_select_autor<-as.list(Data_CR_COD_ANTI[r,5])
             var_select_ano<-as.list(Data_CR_COD_ANTI[r,6])
                        
             for (d in 1:length(r)){
                 
                                     
                data_coc[count,"COD_PROD"]<-var_select_codprod  
                data_coc[count,"PARA"]<-var_select_cod_or[d]
                data_coc[count,"TAG"]<-str_c(var_select_autor[d],var_select_ano[d]) 
                count= count+1                
            } #fecha o for

             
              
      
    } # fecha o for 

end_time <- Sys.time()
end_time - start_time

###################################################
# salvando os dados
###################################################

#arquivos são salvos a partir da data/hora. 
library(lubridate) # biblioteca para manipular datas/horas
tempo<- Sys.time()
tempo<-ymd_hms(tempo)
hora<-hour(tempo)
min<-minute(tempo)
dia<- Sys.Date()
# os arquivos serao salvos em xls
nomearqxls<-paste0("C:\\Users\\Visitante\\Documents\\mybiblio\\PBL_bibliodata_",dia,"_",hora,"_",min,".xlsx") 


wb <- createWorkbook()
addWorksheet(wb, sheetName = "planilha_original")
addWorksheet(wb, sheetName = "lista_artigos_excluidos")
addWorksheet(wb, sheetName = "lista_artigos_doiduplicado")
addWorksheet(wb, sheetName = "planilha_analisada")
addWorksheet(wb, sheetName = "vc_autor")
addWorksheet(wb, sheetName = "vc_titulo")
addWorksheet(wb, sheetName = "rede_coautoria")

writeData(wb, sheet = "planilha_original", x = data)
writeData(wb, sheet = "lista_artigos_excluidos", x = Data_excluidos)
writeData(wb, sheet = "lista_artigos_doiduplicado", x = data_doi_data_select)
writeData(wb, sheet = "planilha_analisada", x = DataPBL) 
writeData(wb, sheet = "vc_autor", x = DataDePara)
writeData(wb, sheet = "vc_titulo", x = DataTituloVC)
writeData(wb, sheet = "rede_coautoria", x = data_comb)
saveWorkbook(wb, file = nomearqxls)

nomearqxls<-paste0("C:\\Users\\Visitante\\Downloads\\mybiblio\\versao_turma\\PBL_citacao.xlsx") # os arquivos serao salvos em xls


wb <- createWorkbook()
addWorksheet(wb, sheetName = "artigo_rede")
addWorksheet(wb, sheetName = "citacao_vc")
addWorksheet(wb, sheetName = "rede_citacao_completa")
addWorksheet(wb, sheetName = "rede_citacao")
addWorksheet(wb, sheetName = "rede_citacao_criarnet")

writeData(wb, sheet = "artigo_rede", x = data)
writeData(wb, sheet = "citacao_vc", x = DataCitacaovc)
writeData(wb, sheet = "rede_citacao_completa", x = data_cocitacao)
writeData(wb, sheet = "rede_citacao", x = cocitacao)
writeData(wb, sheet = "rede_citacao_criarnet", x = data_coc)
saveWorkbook(wb, file = nomearqxls)
