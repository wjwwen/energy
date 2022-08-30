# Global Power Plants Analysis
# Dataset: https://datasets.wri.org/dataset/globalpowerplantdatabase

# Powerplants include thermal (coal, gas, oil, nuclear, biomass, waste, geothermal)
# and renewables (hydro, wind, solar)

import pandas as pd

plant_df = pd.read_csv('global_power_plant_database.csv')

plant_df

plant_df.columns
plant_df.info()

selected_column = [
    'country','country_long','name','capacity_mw','primary_fuel',
    'other_fuel1','other_fuel2','other_fuel3',
    'commissioning_year','year_of_capacity_data',
    'generation_gwh_2013','generation_gwh_2014','generation_gwh_2015',
    'generation_gwh_2016','generation_gwh_2017','estimated_generation_gwh'
]

len(selected_column)

# %%
# Visualisations
import seaborn as sns
import matplotlib
import matplotlib.pyplot as plt

# unique data - 164 countries
plant_df.country_long.nunique()

# Top 20 countries with amount of power plants
countries_plant = plant_df.country_long.value_counts().head(20)
countries_plant

sns.barplot(x = countries_plant.index, y = countries_plant)
plt.xticks(rotation = 90)
plt.title('Country Designation')
plt.ylabel('Number of Power Plant')
plt.xlabel('Countries');

# %%
# Type of fuel in %
main_primary_fuel = plant_df.primary_fuel.value_counts() * 100 / plant_df.primary_fuel.count()
main_primary_fuel

sns.barplot(x = main_primary_fuel, y = main_primary_fuel.index)
plt.title('Main primary fuel')
plt.xlabel('Count (Percentages)');
plt.ylabel('Fuel');

# %%
# Power plant and capacity
sns.scatterplot(x = plant_df.capacity_mw, y = plant_df.primary_fuel, s = 150)
plt.title('Type of power plant and capacity');

# %%
# Capacity of generating power in top 20 countries
countries_capacity = plant_df.groupby('country_long')[['capacity_mw']].sum().sort_values('capacity_mw', ascending = False).head(20)
countries_capacity

sns.barplot(x = countries_capacity.index, y = countries_capacity.capacity_mw)
plt.xticks(rotation = 90)
plt.title('Countries with capacity');

# %%
# Top 10 countries with renewables and fossil fuel plants
renewable_energy = plant_df[plant_df.primary_fuel.isin(['Hydro', 'Wind', 'Solar', 'Biomass', 'Wave and Tidal', 'Geothermal', 'Storage'])]
number_of_renewable_energy = renewable_energy.country_long.value_counts().head(10)
number_of_renewable_energy

fosil_fuel = plant_df[plant_df.primary_fuel.isin(['Gas', 'Oil', 'Coal', 'Nuclear', 'Petcoke', 'Cogeneration'])]
number_of_fosil_fuel_plant = fosil_fuel.country_long.value_counts().head(10)
number_of_fosil_fuel_plant

# %%
# Type of power plant and capacity in China
china_power_plant = plant_df[plant_df.country_long == 'China']

total_power_plant = china_power_plant.country_long.value_counts()
total_power_plant

total_capacity = china_power_plant.capacity_mw.sum()
print('China has a total capacity {} megawatt.'.format(total_capacity))

sns.scatterplot(x = china_power_plant.primary_fuel, y = china_power_plant.capacity_mw, s = 150)
plt.title('Type of power plant and their capacity')