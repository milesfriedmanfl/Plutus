#!/bin/sh

########################################### EDIT THE VARIABLES HERE: #########################################
                                          
EXISTING_BUDGET_PATH="/file/path/of/ExistingBudget.xlsx"                
NEW_BUDGET_PATH="/desired/file/path/NewBudget.xlsx"                   
                                        
##############################################################################################################

# Copy budget template with starting month january to the directory of NEW_BUDGET_PATH with a new file name provided in NEW_BUDGET_PATH:
cp BudgetTemplates/BudgetTemplate_StartMon_Jan.xlsx $NEW_BUDGET_PATH

# Run program with arguments
echo Creating new budget from existing...
java -jar CreateBudgetFromExisting/target/CreateBudgetFromExisting-1.0-SNAPSHOT-jar-with-dependencies.jar $EXISTING_BUDGET_PATH $NEW_BUDGET_PATH
echo Completed creating new budget from existing!

exit 0
