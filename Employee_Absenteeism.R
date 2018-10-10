setwd("D:\\edwisor\\Employee absenteeism")

# Load require Packages
p <- c("xlsx","DMwR","corrgram","caret","usdm","rpart","DataCombine","randomForest",
       "e1071")
lapply(p, require, character.only=TRUE)

rm(p)

D = read.xlsx("Absenteeism_at_work_Project.xls", sheetIndex = 1, header = T)
head(D)
str(D)
cnames <- colnames(D)

# Check number of unique variables
for (i in cnames){
 print(i)
 print(aggregate(data.frame(count = D[,i]), 
 list(value = D[,i]), length))
}

# Data Preprocessing 
preprocessing <- function(D){
   D$ID <- as.factor(D$ID)
   for (i in (1:nrow(D))){
     if (D$Absenteeism.time.in.hours[i] != 0 || is.na(D$Absenteeism.time.in.hours[i])){
       if(D$Reason.for.absence[i] == 0 || is.na(D$Reason.for.absence[i])){
         D$Reason.for.absence[i] = NA
       }
      
      if(D$Month.of.absence[i] == 0 || is.na(D$Month.of.absence[i])){
         D$Month.of.absence[i] = NA
       }
     }
   }
  
  D$Reason.for.absence <- as.factor(D$Reason.for.absence)
  D$Month.of.absence <- as.factor(D$Month.of.absence)
  D$Day.of.the.week <- as.factor(D$Day.of.the.week)
  D$Seasons <- as.factor(D$Seasons)
  D$Disciplinary.failure <- as.factor(D$Disciplinary.failure)
  D$Education <- as.factor(D$Education)
  D$Son <- as.factor(D$Son)
  D$Social.drinker <- as.factor(D$Social.drinker)
  D$Social.smoker <- as.factor(D$Social.smoker)
  D$Pet <- as.factor(D$Pet)
  return(D)
}

D <- preprocessing(D)

#selecting only factor
get_cat <- function(data) {
  return(colnames(data[,sapply(data, is.factor)]))
}

cat_cnames <- get_cat(D)

#selecting only numeric
get_num <- function(data) {
  return(colnames(data[,sapply(data, is.numeric)]))
}

num_cnames <- get_num(D)

############################### Data PreProcessing #########################

# Missing Value Analysis
# Get Missing Values for all Variables

missingValueCheck <- function(data){
  for (i in colnames(data)){
    print(i)
    print(sum(is.na(data[i])))
  }
  print("Total")
  print(sum(is.na(D)))
}

missingValueCheck(D)

# Impute values related to ID
Depenedent_ID <- c("ID","Transportation.expense","Service.time","Age","Height",
                   "Distance.from.Residence.to.Work","Education","Son","Weight",
                   "Social.smoker","Social.drinker","Pet","Body.mass.index")
Depenedent_ID_data <- D[,Depenedent_ID]
Depenedent_ID_data <- aggregate(. ~ ID, data = Depenedent_ID_data, 
                                FUN = function(e) c(x = mean(e)))

for (i in Depenedent_ID) {
  for (j in (1:nrow(D))){
    ID <- D[j,"ID"]
    if(is.na(D[j,i])){
      D[j,i] <- Depenedent_ID_data[ID,i]
    }
  }
}

# Covert factor variable values to labels 

for (i in cat_cnames){
  D[,i] <- factor(D[,i], 
  labels = (1:length(levels(factor(D[,i])))))
}

# Impute values for other variables
library(mice)
imputed_data <- mice(D, m=1, maxit = 50, method = "cart", seed = 500)
D <- complete(imputed_data, 1)

#D = knnImputation(D, k = 7)
#missingValueCheck(D_imp)

missingValueCheck(D)


# Outlier Analysis
# Histogram

for(i in num_cnames){
  hist(D[,i], xlab=i, main=" ", col=(c("lightblue","darkgreen")))
}


# BoxPlot
num_cnames <- num_cnames[ num_cnames != "Absenteeism.time.in.hours"]
for(i in num_cnames){
  boxplot(D[,i]~D$Absenteeism.time.in.hours,
  data=D, main=" ", ylab=i, xlab="Absenteeism.time.in.hours",  
  col=(c("lightblue","darkgreen")), outcol="red")
}

# Removing ID related varaibles from num_cnames

for (i in Depenedent_ID){
  num_cnames <- num_cnames[ num_cnames != i]
}

# Replace all outliers with NA and impute

for(i in num_cnames){
  val = D[,i][D[,i] %in% boxplot.stats(D[,i])$out]
  D[,i][D[,i] %in% val] = NA
}

# Impute NA Values

missingValueCheck(D)

#D <- knnImputation(D, k = 7)

imputed_data <- mice(D, m=1, maxit = 50, method = "cart", seed = 500)
D <- complete(imputed_data, 1)
missingValueCheck(D)

# Copy Employee Data into DataSet for further analysis

DataSet <- D

######################## Feature Selection #################

# Correlation Plot
corrgram(D, upper.panel=panel.pie, text.panel=panel.txt, 
         main = "Correlation Plot")


# ANOVA
cnames <- colnames(D)

for (i in cnames){
  print(i)
  print(summary(aov(D$Absenteeism.time.in.hours~D[,i], D)))
}

# Dimensionality Reduction
D <- subset(D, select = -c(Weight, Education, Service.time, Social.smoker,Body.mass.index,
                           Work.load.Average.day.,Seasons, Transportation.expense,Pet,
                           Disciplinary.failure, Hit.target, Month.of.absence,Social.drinker))



######################### Fature Normalization ######################

cat_cnames <- get_cat(D)
num_cnames <- get_num(D)

# Fature Scaling

for (i in num_cnames){
  D[,i] <- (D[,i] - min(D[,i])) / (max(D[,i]) - min(D[,i]))
}

#################### Model Development ########################

# Data Divide

X_index <- sample(1:nrow(D), 0.8 * nrow(D))
X_train <- D[X_index,-8]
X_test <- D[-X_index,-8]
y_train <- D[X_index,8]
y_test <- D[-X_index,8]
train <- D[X_index,]
test <- D[-X_index,]

#calculate RMSE
RMSE <- function(y, yhat){
  sqrt(mean((y - yhat)^2))
}

#calculate MSE
MSE <- function(y, yhat){
  (mean((y - yhat)^2))
}

###################### Multiple Linear Regression ########################
num_d <- train[sapply(train, is.numeric)]
vifcor(num_d, th=0.9)

lm_regressor <- lm(Absenteeism.time.in.hours~.,data = train)

summary(lm_regressor)

#Predict for new test cases

for (i in cat_cnames){
  
  lm_regressor$xlevels[[i]] <- union(lm_regressor$xlevels[[i]], 
                                     
                                     levels(X_test[[i]]))
  
}

lm_predict <- predict(lm_regressor, newdata=X_test)
RMSE(lm_predict,y_test)
MSE(lm_predict,y_test)

##############################Decision Trees for regression ###################
DT_regressor <- rpart(Absenteeism.time.in.hours~.,data = train, method="anova")

#Predict for new test cases

DT_predict <- predict(DT_regressor, X_test)

RMSE(DT_predict,y_test)
MSE(DT_predict,y_test)

##################### Random Forest ########################

RF_regressor <- randomForest(x = X_train, y = y_train, ntree = 100)

#Predict for new test cases

RF_predict <- predict(RF_regressor, X_test)
RMSE(RF_predict,y_test)
MSE(RF_predict,y_test)

#################### Support Vector Regressor #####################
SVR_regressor <- svm(formula = Absenteeism.time.in.hours ~ ., 
                     data = train, type = 'eps-regression')

#Predict for new test cases

SVR_predict <- predict(SVR_regressor, X_test)
RMSE(SVR_predict,y_test)
MSE(SVR_predict,y_test)

############################## Problems ##################################

# Suggesting the changes

lm_regressor_p1 <- lm(Absenteeism.time.in.hours~.,data = D)
summary(lm_regressor_p1)

# Calculating Losses
p2_data = D[,-8]

#Predict for new test cases
p2_predict <- predict(SVR_regressor, p2_data)

# Convert predict values back to original scale
p2_predict <- (p2_predict * 120)

# Add predicted values to the DataSet
p2_dataSet <- merge(DataSet,p2_predict,by="row.names",all.x=TRUE)

# Calculate the total Loss 
Loss <- 0

for (i in 1:nrow(p2_dataSet)){
  
  if (p2_dataSet$Hit.target[i] != 100)
    
    if (p2_dataSet$Age[i] >= 25 &&  p2_dataSet$Age[i] <= 32){
      
      Loss = Loss + as.numeric(p2_dataSet$Disciplinary.failure[i]) * 2000 +
        
        (as.numeric(p2_dataSet$Education[i]) + 1) * 500 + p2_dataSet$y[i] * 1000
      
    }else if(p2_dataSet$Age[i] >= 33 &&  p2_dataSet$Age[i] <= 40){
      
      Loss = Loss + as.numeric(p2_dataSet$Disciplinary.failure[i]) * 2000 +
        
        (as.numeric(p2_dataSet$Education[i]) + 2) * 500 + p2_dataSet$y[i] * 1000
      
    }else if(p2_dataSet$Age[i] >= 41 &&  p2_dataSet$Age[i] <= 49){
      
      Loss = Loss + as.numeric(p2_dataSet$Disciplinary.failure[i]) * 2000 +
        
        (as.numeric(p2_dataSet$Education[i]) + 3) * 500 + p2_dataSet$y[i] * 1000
      
    }else if(p2_dataSet$Age[i] >= 50 &&  p2_dataSet$Age[i] <= 60){
      
      Loss = Loss + as.numeric(p2_dataSet$Disciplinary.failure[i]) * 2000 +
        
        (as.numeric(p2_dataSet$Education[i]) + 4) * 500 + p2_dataSet$y[i] * 1000
      
    } 
  
}

# To calculate loss per month

Loss <- Loss/12

Loss
