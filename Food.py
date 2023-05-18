import xlrd
#open workbook
workbook = xlrd.open_workbook("Macros.xlsx")
#open worksheet
worksheet = workbook.sheet_by_index(0) #has one sheet

# foods and components sorted alphabetically
# spaces included
# list of components, list of foods
# take list of foods and components from excel
class Food(object):
    
    name = ""  # accepts spaces
    components = []  # list of components in food
    cals = 0  # per serving
    carbs = 0
    protein = 0
    fat = 0
    price = 0  # total of prices per servings
    # comprised of multiple serving sizes

    def __init__(self, name, components):
        self.name = name
        self.components = components  # transfer components data
        for x in range(len(components)):
            self.cals = self.cals + components[x].cals
            self.carbs = self.carbs + components[x].carbs
            self.protein = self.protein + components[x].protein
            self.fat = self.fat + components[x].fat
            self.price = self.price + components[x].price

    def set_name(self, new_name):
        self.name = new_name
        
    def set_cals(self, new_cals):
        self.cals = new_cals

    def set_carbs(self, new_carbs):
        self.carbs = new_carbs

    def set_protein(self, new_protein):
        self.protein = new_protein

    def set_fat(self, new_fat):
        self.fat = new_fat

    def set_serving(self, new_serving):
        self.serving_size = new_serving

    def set_price(self, new_price):
        self.price = new_price
        
    def get_name(self):
        return self.name

    def get_cals(self):
        return self.cals

    def get_carbs(self):
        return self.carbs

    def get_protein(self):
        return self.protein

    def get_fat(self):
        return self.fat

    def get_serving(self):
        return self.serving_size

    def get_price(self):
        return self.price  

            
class Component(object):

    __name = ""
    cals = 0  # per serving
    carbs = 0
    protein = 0
    fat = 0
    serving_size = 0  # grams
    #price = 0  # price per serving
      
    def __init__(self, name, cals, carbs, protein, fat, serving_size, ):
        self.name = name
        self.cals = cals
        self.carbs = carbs
        self.protein = protein
        self.fat = fat
        self.serving_size = serving_size
        #self.price = price
        
    def set_name(self, new_name):
        self.name = new_name
        
    def set_cals(self, new_cals):
        self.cals = new_cals

    def set_carbs(self, new_carbs):
        self.carbs = new_carbs

    def set_protein(self, new_protein):
        self.protein = new_protein

    def set_fat(self, new_fat):
        self.fat = new_fat

    def set_serving(self, new_serving):
        self.serving_size = new_serving
"""
    def set_price(self, new_price):
        self.price = new_price
    """    
    def get_name(self):
        return self.name

    def get_cals(self):
        return self.cals

    def get_carbs(self):
        return self.carbs

    def get_protein(self):
        return self.protein

    def get_fat(self):
        return self.fat

    def get_serving(self):
        return self.serving_size
    """
    def get_price(self):
        return self.price
"""


