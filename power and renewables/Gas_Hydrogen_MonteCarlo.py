"""
<< Financial risk analysis/Monte Carlo Simulation - Conversion of mature gas field to blue hydrogen project >>

A Monte Carlo simulation is a model used to predict the probability of different outcomes 
when the intervention of random variables is present. 
Monte Carlo simulations help to explain the impact of risk and 
uncertainty in prediction and forecasting models.

Calculate NPV for both gas and blue hydrogen projects.

Assume same revenues for both projects over 5 years, discount factor = 10%, and that only OPEX matters. 
The operator knows with certainty that OPEX for gas extraction is $5/unit.

Free cash flow = revenue - expenses
NPV = future cash flow of a project in present-day terms using time value of money

"""

import numpy as np  
import numpy.random as npr  
import matplotlib.pyplot as plt  

# ------------------------ NPV of Gas project
revenue_gas = np.array([10, 10, 10, 10, 10])
opex_gas = np.array([5, 5, 5, 5, 5])
discount_factor = 0.1

discounted_cashflow = np.empty((0, 0))  # Container for npv

# Discounted cash flow model
for t in np.arange(0, len(revenue_gas)):
    net_profits = revenue_gas[t] - opex_gas[t]
    nfc = (net_profits) / (1 + discount_factor) ** t
    discounted_cashflow = np.append(discounted_cashflow, nfc)

npv_gas = round(discounted_cashflow.sum(), 2)  # Round to two decimals

# ------------------------ NPV of the Blue Hydrogen
# Simulate OPEX 
npr.seed(12634)  # Random seed so example is reproducible
n_periods = 5  # No. of simulated years
n_sims = 1000  # No. of simulated cost paths
initial_cost = 6.5
drift = -0.35  # Trend for cost path
mu = 0  # Parameter for random number generator
sigma = 0.3  # Parameter for random number generator

# Function for simulation reproducibility with different values for the assumptions
def simulate_costs(n_periods, n_sims, initial_cost, drift, mu, sigma):
    simulations = np.empty((1, 5))  # results container
    for sim in np.arange(1, n_sims):
        simulated_uncertainty = np.cumsum(npr.normal(mu, sigma, n_periods))
        time = np.arange(1, n_periods + 1)
        simulated_cost = initial_cost + time * drift + simulated_uncertainty
        simulations = np.concatenate((simulations, [simulated_cost]))

    # Remove the first member of the array that came with the initializer
    simulations = simulations[1:]
    return simulations

simulated_costs = simulate_costs(n_periods, n_sims, initial_cost, drift, mu, sigma)

# ------------------------ Calculate NPV for each simulated cost path
revenue_hydrogen = np.array([10, 10, 10, 10, 10])
discount_factor = 0.1
npv_hydrogen = np.empty((0, 0))  # Container for NPV
opex = simulated_costs[875]

# NPV function
def calculate_hydrogen_npv(rev_hydrogen, opex_sim, discount_factor_sim):
    hydrogen_discounted_cashflow = np.empty((0, 0))  # Container for net free cashflow

    for t in np.arange(0, len(rev_hydrogen)):
        net_profits = rev_hydrogen[t] - opex_sim[t]
        nfc = (net_profits) / (1 + discount_factor_sim) ** t
        hydrogen_discounted_cashflow = np.append(hydrogen_discounted_cashflow, nfc)

    npv = round(hydrogen_discounted_cashflow.sum(), 2)  # Round to two decimals
    return npv

# ------------------------ NPV for all simulated paths
for i in np.arange(0, np.shape(simulated_costs)[0]):
    opex = simulated_costs[i]
    x = calculate_hydrogen_npv(revenue_hydrogen, opex, discount_factor)
    npv_hydrogen = np.append(npv_hydrogen,x)

mean_npv_hydrogen = np.median(npv_hydrogen)

# ------------------------ Plot results
# Initialize plotting object
fig, axs = plt.subplots(nrows=1, ncols=1, figsize=(6, 4), facecolor='w', edgecolor='k')

# Histogram to check distribution of Hydrogen NPVs
axs.hist(npv_hydrogen, bins=100, density=True, color='#1A5276', alpha=.5)

# Plot contrast lines for gas and hydrogen projects
axs.axvline(x=npv_gas, color='#6C3483', linestyle='solid', label = 'NPV of gas project')
axs.axvline(x=mean_npv_hydrogen, color='red', linestyle='solid', label = 'Expected NPV of hydrogen')

# Plot
plt.legend()
plt.title('Distribution of NPV values of hydrogen project')
plt.ylabel('Probability density of each NPV value')
plt.xlabel('Net present value simulation in £')
plt.xticks()

plt.show()
fig.savefig('NPV_distribution.jpeg')

# ------------------------ Plot simulated cost paths
fig, axs = plt.subplots(nrows=1, ncols=1, figsize=(6, 4), facecolor='w', edgecolor='k')
for i in np.arange(0, np.shape(simulated_costs)[0]):
    axs.plot(simulated_costs[i], color='#CACFD2')

plt.axhline(initial_cost, color='red', label='Assumed intial hydrogen opex')
plt.axhline(5, color='#6C3483', label='Assumed gas opex')

plt.legend()
plt.title('Simulated cost paths from a random walk with drift')
plt.ylabel('Hydrogen Opex in £')
plt.xlabel('Year')
plt.xticks([0, 1, 2, 3,4], [1,2,3,4,5])

plt.show()
fig.savefig('simulated_costs_paths.jpeg')