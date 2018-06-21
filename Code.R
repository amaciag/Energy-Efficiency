# Import appropriate libraries
library(ClustOfVar)
library(PCAmixdata)
library(FactoMineR)
library(rpart)
library(caret)
library(splitstackshape)
library(readxl)
library(tidyverse)

# Read in file, select worksheet and convert to dataframe
# Some of the column names are changed automatically
# such as from EnergyConsumption - kWh to Energy.Consumption...kWh
filename="EnergyData - sample.xlsx"
data<-read_excel(filename,sheet='Energy')
data<-data.frame(data)

# Create a new variable 'month' from variable PeriodLabel
# Split Energy Consumption separately for clustering purposes
# We will cluster buildings' characteristics
# Will also create a unique id for each row in dataset by
# combining integer rownames and Site in order to identify individual data
data$month<-as.numeric(format(data$PeriodLabel,"%m"))
targets<-data$Energy.Consumption...kWh
data$index<-paste(rownames(data),data$SiteCode,sep='-')
rownames(data)<-data$index
data$index<-NULL

# Drop unnecessary variables
data<-subset(data,select=-c(Energy.Consumption...kWh,SiteCode,PropertyType,isDeleted,
                            JLLResp,PeriodLabel,Country,City,Address,Property.Name))


# Distinguish between numerical and categorical variables
X.quanti<-subset(data,select=c(EnergyCost....,SqFtNIA,SqFtRentable,Latitude,Longitude))
X.quali<-subset(data,select=-c(EnergyCost....,SqFtNIA,SqFtRentable,Latitude,Longitude))

# Convert to numerical type
for(each in colnames(X.quanti)){
  unlisting<-unlist(X.quanti[each])
  X.quanti[each]<-as.numeric(as.character(unlisting))
}

# Convert to categorical type
for(each in colnames(X.quali)){
  unlisting<-unlist(X.quali[each])
  X.quali[each]<-as.factor(unlisting)
}

# Combine numerical and categorical variables 
# Binarize categorical data
# Drop NULL and N/A variables
X=cbind(X.quanti,tab.disjonctif(X.quali))
X['NULL']<-NULL
X['N/A']<-NULL


# Perform hierarchical clustering
# Draw a dendrogram
tree<-hclustvar(X)
plot(tree)

# Perform stability transformation to decide on number of clusters
# Plot results
stab<-stability(tree,B=20)
plot(stab,main='Stability of partitions')
boxplot(stab$matCR,main='Dispersion of the adjusted Rand index')

# Cut the tree from 18-46 clusters
P46<-cutreevar(tree,46)
P44<-cutreevar(tree,44)
P42<-cutreevar(tree,42)
P40<-cutreevar(tree,40)
P38<-cutreevar(tree,38)
P36<-cutreevar(tree,36)
P34<-cutreevar(tree,34)
P32<-cutreevar(tree,32)
P30<-cutreevar(tree,30)
P28<-cutreevar(tree,28)
P26<-cutreevar(tree,26)
P24<-cutreevar(tree,24)
P22<-cutreevar(tree,22)
P20<-cutreevar(tree,20)
P18<-cutreevar(tree,18)

# Based on stability of partitions, clusters from 18 to 46.
# More clusters would single out and turn into single clusters
# I want each variable to be pair up, so 28 clusters seem legitimate
size_cluster46<-data.frame(P46$size)
size_cluster30<-data.frame(P30$size)
size_cluster28<-data.frame(P28$size)
size_cluster26<-data.frame(P26$size)


# Gain in cohesion for each
clu<-c(18,20,22,24,26,28,30,32,34,36,38,40,42,44,46)
hier<-c(P18$E,P20$E,P22$E,P24$E,P26$E,P28$E,P30$E,P32$E,P34$E,P36$E,P38$E,P40$E,P42$E,P44$E,P46$E)
gain<-do.call(rbind,Map(data.frame, Cluster=clu, Gain_in_Cohesion=hier))

# Choose 28 clusters and assign classes to data
# Add variable Energy Consumption back to data with assigned clusters
hlabels<-data.frame(P28$scores)
output<-X
output$EnergyConsumption<-as.numeric(as.character(unlist(targets)))
output$hcluster<-colnames(hlabels)[apply(hlabels,1,which.max)]


# Create a benchmark column to store average of predicted values from decision tree regression
# Create a dataframe to store metrics from decision tree for each cluster
output$benchmark<-' '
tree_metrics<-data.frame(matrix(nrow=0,ncol=3))


# For each cluster, build a cluster model formed by decision tree
# For each decision tree, predict values
# Take average of the predicted values and store it as benchmark
# Save the metrics results for each tree
unique_clusters<-unique(output$hcluster)
for (each in unique_clusters){
  cluster<-subset(output,hcluster==each)
  num_cluster<-as.numeric(substring(each,8))
  clustera<-cluster[,which(P28$cluster==num_cluster)]
  clustera$EnergyCon<-cluster$EnergyConsumption
  reg<-rpart(formula=EnergyCon~.,data=clustera,method='anova',control=rpart.control(minsplit=5))
  y_p<-predict(reg,clustera)
  benchmark<-mean(y_p)
  clustera$prediction<-y_p
  clustera$benchmark<-benchmark
  output[which(output$hcluster==each),]$benchmark<-clustera$benchmark
  metrics<-postResample(pred=y_p,obs=clustera$EnergyCon)
  tree_metrics<-rbind(tree_metrics,metrics)
}

# Tree metrics results
names(tree_metrics)<-c("RMSE",'Rsquared','MAE')
tree_metrics$Cluster<-unique_clusters

# Label data as inefficient if actual energy consumption is greater than benchmark
output$benchmark<-as.numeric(output$benchmark)
output$label<-ifelse(output$EnergyConsumption>output$benchmark,'inefficient','efficient')


# Get SiteCode variable back
output$index<-rownames(output)
output<-cSplit(output,"index","-")
output$index_1<-NULL
colnames(output)[which(names(output)=='index_2')]<-'SiteCode'

# Save results in csv file
write.csv(output,file='output.csv')





