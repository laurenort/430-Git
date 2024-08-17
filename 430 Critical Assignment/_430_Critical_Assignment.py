# Importing the necessary modules
import math
import pandas as pd

# Defining the main function call
def main():       
    # Performing calculations(min, max, mean, std) for the numerical values(Items, Net Sales, Age) in the dataset.
    
    # Using pandas module and read_excel method to open the file and the specified sheet name while selecting the specific
    # columns(Items, Net Sales, and Age) and storing inside a variable.
    # Then taking said variable and using the DataFrame method from pandas and storing inside a variable.
    file = pd.read_excel('C:/Users/lauren/Downloads/Dataset.xlsx', sheet_name = 'Data', usecols = "C, D, H")  
    df = pd.DataFrame(file)
    
    # Here I am extracting specific columns from the file and storing them to proper variables.
    items = df['Items']
    net_sales = df['Net Sales']
    age = df['Age']

    # Printing the specified category along with the calculations.
    # Also rounding up the mean and std to 2 digits for cleaner output.
    # Using several methods from the math module to perform calculations.
    # This is only for the 'Items' column.
    print("Numerical Values")
    print("********************")
    print("(Items)")
    print("Minimum: ", items.min())
    print("Max: ", items.max())
    print("Mean: ", round(items.mean(), 2))
    print("Standard Deviation: ", round(items.std(), 2))

    # Same thing but for the 'Net Sales' column.
    print()
    print("(Net Sales)")
    print("Minimum: ", net_sales.min())
    print("Max: ", net_sales.max())
    print("Mean: ", round(net_sales.mean(), 2))
    print("Standard Deviation: ", round(net_sales.std(), 2))

    # Same thing but for the 'Age' column.
    print()
    print("(Age)")
    print("Minimum: ", age.min())
    print("Max: ", age.max())
    print("Mean: ", round(age.mean(), 2))
    print("Standard Deviation: ", round(age.std(), 2))
    

    # Calculating the categorical percentages of:
        # Customer
        #   (Regular)
        #   (Promotional)
        # Gender
        #   (Male)
        #   (Female)
        # Method of Payment
        #   (Discover)
        #   (Proprietary Card)
        #   (MasterCard)
        #   (Visa)
        # Martial Status
        #   (Married)
        #   (Single)
    
    # Using pandas module and read_excel method to open the file and the specified sheet name while selecting a specific
    # column(Customer) and skipping the title row and storing inside a variable.
    # Then taking said variable and using the DataFrame method from pandas and storing inside a variable.
    a = pd.read_excel('C:/Users/lauren/Downloads/Dataset.xlsx', sheet_name = 'Data', usecols = "B", skiprows = 1)  
    df = pd.DataFrame(a)   

    # Here I am defining the item values that are in the column.
    CUS_REG = "Regular"
    CUS_PRO = "Promotional"
    
    # Using the value_counts method from DataFrame so that it counts the number of
    # unique values and storing inside variable count.
    count = df.value_counts()
    
    # Defining two counters for regular and promotional customers.
    # Using the get method so that it returns the item of the specified key and
    # returns 0 by default and storing inside separate variables.
    counter1 = count.get(CUS_REG, 0)
    counter2 = count.get(CUS_PRO, 0)
    
    # Using the size property from DataFrame that will return the number of elements
    # and storing inside variable total.
    total = df.size
    
    # Performing calculation to get the percentage of regular and promotional customers.
    # It takes the number of the specific item and divides by the total items and multiplies
    # by 100 to get the percentage, then it is stores inside a variable.
    percent1 = (counter1 / total) * 100
    percent2 = (counter2 / total) * 100
    
    # Printing the specified category along with the calculations.
    # Rounding up the percentages for cleaner output.
    # This is for regular and promotional customers.
    print()
    print("Categorical Percentages")
    print("********************")
    print("Regular Customers: ", round(percent1), "%")
    print("Promotional Customers: ", round(percent2), "%")
    

    # Using pandas module and read_excel method to open the file and the specified sheet name while selecting a specific
    # column(Gender) and skipping the title row and storing inside a variable.
    # Then taking said variable and using the DataFrame method from pandas and storing inside a variable.
    b = pd.read_excel('C:/Users/lauren/Downloads/Dataset.xlsx', sheet_name = 'Data', usecols = "F", skiprows = 1)  
    df = pd.DataFrame(b)   

    # Here I am defining the item values that are in the column.
    GEN_MALE = "Male"
    GEN_FEMALE = "Female"
    
    # Using the value_counts method from DataFrame so that it counts the number of
    # unique values and storing inside variable count.
    count = df.value_counts()
    
    # Defining two counters for male and female gender.
    # Using the get method so that it returns the item of the specified key and
    # returns 0 by default and storing inside separate variables.
    counter1 = count.get(GEN_MALE, 0)
    counter2 = count.get(GEN_FEMALE, 0)
    
    # Using the size property from DataFrame that will return the number of elements
    # and storing inside variable total.
    total = df.size
    
    # Performing calculation to get the percentage of male and female gender.
    # It takes the number of the specific item and divides by the total items and multiplies
    # by 100 to get the percentage, then it is stores inside a variable.
    percent1 = (counter1 / total) * 100
    percent2 = (counter2 / total) * 100
    
    # Printing the calculations still under the same category.
    # Rounding up the percentages for cleaner output.
    # This is for male and female gender.
    print()
    print("Male Gender: ", round(percent1), "%")
    print("Female Gender: ", round(percent2), "%")
    

    # Using pandas module and read_excel method to open the file and the specified sheet name while selecting a specific
    # column(Method of Payment) and skipping the title row and storing inside a variable.
    # Then taking said variable and using the DataFrame method from pandas and storing inside a variable.
    c = pd.read_excel('C:/Users/lauren/Downloads/Dataset.xlsx', sheet_name = 'Data', usecols = "E", skiprows = 1)  
    df = pd.DataFrame(c)   

    # Here I am defining the item values that are in the column.
    PAY_DIS = "Discover"
    PAY_PCARD = "Proprietary Card"
    PAY_MCARD = "MasterCard"
    PAY_VISA = "Visa"
    
    # Using the value_counts method from DataFrame so that it counts the number of
    # unique values and storing inside variable count.
    count = df.value_counts()
    
    # Defining four counters for all the different payment types.
    # Using the get method so that it returns the item of the specified key and
    # returns 0 by default and storing inside separate variables.
    counter1 = count.get(PAY_DIS, 0)
    counter2 = count.get(PAY_PCARD, 0)
    counter3 = count.get(PAY_MCARD, 0)
    counter4 = count.get(PAY_VISA, 0)
    
    # Using the size property from DataFrame that will return the number of elements
    # and storing inside variable total.
    total = df.size
    
    # Performing calculation to get the percentage of all the different payment types.
    # It takes the number of the specific item and divides by the total items and multiplies
    # by 100 to get the percentage, then it is stores inside a variable.
    percent1 = (counter1 / total) * 100
    percent2 = (counter2 / total) * 100
    percent3 = (counter3 / total) * 100
    percent4 = (counter4 / total) * 100
    
    # Printing the calculations still under the same category.
    # Rounding up the percentages for cleaner output.
    # This is for the different payment types.
    print()
    print("Discover Method: ", round(percent1), "%")
    print("Proprietary Card Method: ", round(percent2), "%")
    print("Master Card Method: ", round(percent3), "%")
    print("Visa Method: ", round(percent4), "%")


    # Using pandas module and read_excel method to open the file and the specified sheet name while selecting a specific
    # column(Marital Status) and skipping the title row and storing inside a variable.
    # Then taking said variable and using the DataFrame method from pandas and storing inside a variable.
    d = pd.read_excel('C:/Users/lauren/Downloads/Dataset.xlsx', sheet_name = 'Data', usecols = "G", skiprows = 1)  
    df = pd.DataFrame(d)   

    # Here I am defining the item values that are in the column.
    STAT_MAR = "Married"
    STAT_SIN = "Single"
    
    # Using the value_counts method from DataFrame so that it counts the number of
    # unique values and storing inside variable count.
    count = df.value_counts()
    
    # Defining two counters for married and single status.
    # Using the get method so that it returns the item of the specified key and
    # returns 0 by default and storing inside separate variables.
    counter1 = count.get(STAT_MAR, 0)
    counter2 = count.get(STAT_SIN, 0)
    
    # Using the size property from DataFrame that will return the number of elements
    # and storing inside variable total.
    total = df.size
    
    # Performing calculation to get the percentage of the two marital status.
    # It takes the number of the specific item and divides by the total items and multiplies
    # by 100 to get the percentage, then it is stores inside a variable.
    percent1 = (counter1 / total) * 100
    percent2 = (counter2 / total) * 100
    
    # Printing the calculations still under the same category.
    # Rounding up the percentages for cleaner output.
    # This is for married and single status.
    print()
    print("Married Status: ", round(percent1), "%")
    print("Single Status: ", round(percent2), "%")
    
    #   Customer Values:
    #       - Items by Customer(Regular & Promotional)
    #           - Min
    #           - Max
    #           - Mean
    #           - Standard Deviation
    #       - Net Sales by Customer(Regular & Promotional)
    #           - Min
    #           - Max
    #           - Mean
    #           - Standard Deviation
    
    # Using pandas module and read_excel method to open the file and the specified sheet name while selecting the specific
    # columns(Customer and Items) and storing inside a variable.
    # Then taking said variable and using the DataFrame method from pandas and storing inside a variable.
    e = pd.read_excel('C:/Users/lauren/Downloads/Dataset.xlsx', sheet_name = 'Data', usecols = "B, C")  
    df = pd.DataFrame(e)
    
    # Using the groupby method from DataFrame that groups data based on the specified rows and columns.
    # Choosing Customer and then Items which is a numerical column, and storing inside a variable.
    group = df.groupby('Type of Customer')['Items']    

    # Using the get_group method from DataFrame that gets the defined item("Regular" and "Promotional")
    # from the group and storing inside specified variable.
    reg = group.get_group('Regular')
    promo = group.get_group('Promotional')

    ### COMMENT ###
    # ITEMS
    print()
    print("Customer Values")
    print("********************")
    print("(Regular Customers - Items)")
    print("Minimum of Items", reg.min())
    print("Max of Items: ", reg.max())
    print("Mean of Items: ", round(reg.mean(), 2))
    print("Standard Deviation of Items: ", round(reg.std(), 2))

    ### COMMENT ###
    # ITEMS
    print()
    print("(Promotional Customers - Items)")
    print("Minimum of Items", promo.min())
    print("Max of Items: ", promo.max())
    print("Mean of Items: ", round(promo.mean(), 2))
    print("Standard Deviation of Items: ", round(promo.std(), 2))

    ### COMMENT ###
    f = pd.read_excel('C:/Users/lauren/Downloads/Dataset.xlsx', sheet_name = 'Data', usecols = "B, D")  
    df = pd.DataFrame(f)
    
    ### COMMENT ###
    group = df.groupby('Type of Customer')['Net Sales']    

    ### COMMENT ###
    reg = group.get_group('Regular')
    promo = group.get_group('Promotional')
    
    ### COMMENT ###
    # NET SALES
    print()
    print("(Regular Customers - Net Sales)")
    print("Minimum of Net Sales", reg.min())
    print("Max of Net Sales: ", reg.max())
    print("Mean of Net Sales: ", round(reg.mean(), 2))
    print("Standard Deviation of Net Sales: ", round(reg.std(), 2))

    ### COMMENT ###
    # NET SALES
    print()
    print("(Promotional Customers - Net Sales)")
    print("Minimum of Net Sales", promo.min())
    print("Max of Net Sales: ", promo.max())
    print("Mean of Net Sales: ", round(promo.mean(), 2))
    print("Standard Deviation of Net Sales: ", round(promo.std(), 2))
        

### COMMENT ###
if __name__ == "__main__" : 
    main()

