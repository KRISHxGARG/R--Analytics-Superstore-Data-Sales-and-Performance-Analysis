library("readxl")             # To Read the Sample Superstore Excel Workbook
Dataset=read_excel("C:/Users/HP/OneDrive/Desktop/MABS/Trimester-2/Business Analytics-2/Assignment and Project/Assignment-4/Superstore Dataset  .xlsx",sheet=1)
library(dplyr)                # Using dplyr for Data Manipulation and Transformation

#To view the Dataset
View(Dataset)                 # To view the complete Dataset

#To show the structure of Dataset
str(Dataset)                  # To view the Structure of Dataset in Console Area

#To view the Top 20 rows of the Dataset
head(Dataset,20)              # To view top 20 rows of the Dataset in Console Area
View(head(Dataset,20))        # To view top 20 rows of the Dataset

#To view the bottom 6 rows of the Dataset
tail(Dataset)                 # To view Bottom 6 rows of the Dataset in Console Area
View(tail(Dataset))           # To view Bottom 6 rows of the Dataset 

#To show the statistical summary of the Dataset
summary(Dataset)              # To view Summary of the Dataset in Console Area 
View(summary(Dataset))        # To view Summary of the Dataset

#To select specific columns
selected_data=Dataset%>%
  select(`Ship Mode`,`Customer Name`)   
View(selected_data)           # To View the Ship Mode and Customer Name in Dataset

#How to Rename the Ship Mode and Postal Code Column?
renamed_data=Dataset%>%
  rename(
    Mode=`Ship Mode`,
    Pin_Code=`Postal Code`
  )
View(renamed_data)            # To View the Dataset after renaming the Ship Mode and Postal Code Columns

#Q.1) How can we view the top 6 rows of the renamed dataset?
View(head(renamed_data))      # To View the Top 6 rows of the Renamed Dataset    

#Q.2) How can we Check for missing values?
colSums(is.na(Dataset))       # To Check the Missing value in the Dataset
View(colSums(is.na(Dataset))) # To View the Missing value in the Dataset

#Q.3) How can we Filtered the Rows based as per our requirements?
filtered_data=Dataset%>%
  filter(`Ship Mode`== "First Class",Segment=="Corporate",State=="Texas")
View(filtered_data)           # To View the Filter data that contains value having Ship Mode= First Class, Segment=Corporate, State= Texas 

#Q.4) How can we view the top 6 rows of the Filtered dataset?
View(head(filtered_data))     # To View the Top 6 rows of the Filtered Dataset

#Q.5) How can we  view the Summary of Category, Sub-Category,Segment,Total Sales and Total Profit?
Short_Summary=Dataset%>%
group_by(Category, `Sub-Category`,Segment) %>%
summarise(Total_Sales=sum(Sales, na.rm = TRUE),
          Total_Quantity=sum(Quantity, na.rm=TRUE),
          Total_Profit=sum(Profit, na.rm = TRUE))
View(Short_Summary)           # To View the Short Summary of the Dataset

#Q.6) How can we find the most profitable category?
most_profitable_category=Dataset%>%
  group_by(Category)%>%
  summarise(Total_Profit = sum(Profit, na.rm = TRUE))%>%
  arrange(desc(Total_Profit)) %>%
  head(1)
View(most_profitable_category)    # To View the Most Profitable Category

#Q.7) How can we find the least profitable category?
least_profitable_category=Dataset%>%
  group_by(Category)%>%
  summarise(Total_Profit = sum(Profit, na.rm = TRUE))%>%
  arrange(Total_Profit) %>%
  head(1)
View(least_profitable_category)    # To View the Least Profitable Category

#Q.8) How to create a new Column 'Unit Price of a Product' by using Mutating and Transforming Data?
mutated_data=Short_Summary%>%
  mutate(
    Unit_Price=Total_Sales/Total_Quantity
  )
View(mutated_data)            # To View the new column of the Dataset
head(mutated_data)           
View(head(mutated_data))      # To View the top 6 rows of the Dataset

#Q.9) How can we find the most profitable Sub-Category?
most_profitable_sub_category=Dataset%>%
  group_by(`Sub-Category`)%>%
  summarise(Total_Profit = sum(Profit, na.rm = TRUE))%>%
  arrange(desc(Total_Profit)) %>%
  head(1)
View(most_profitable_sub_category)    # To View the Most Profitable Sub-Category

#Q.10) How can we find the least profitable Sub-Category?
least_profitable_sub_category=Dataset%>%
  group_by(`Sub-Category`)%>%
  summarise(Total_Profit = sum(Profit, na.rm = TRUE))%>%
  arrange(Total_Profit) %>%
  head(1)
View(least_profitable_sub_category)    # To View the Least Profitable Sub-Category