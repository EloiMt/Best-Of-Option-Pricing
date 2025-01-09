# Best-Of Option Pricing

## **Project Overview**
This project aims to implement and analyze the pricing of two financial derivatives: the **Best-Of Call Option** and the **Worst-Of Call Option**, using real-world data and simulations.

The key steps in this project include:
1. Retrieving the last spot prices of two assets (Société Générale and BNP) from **Boursorama**.
2. Simulating N price series over T time steps for these two correlated assets using a **Geometric Brownian Motion (GBM)** model.
3. Calculating the prices of the Best-Of and Worst-Of options.
4. Comparing the two prices and analyzing the results.

---

## **Files in Repository**
1. **`Best-Of Option pricing.pdf`**  
   - Detailed report of the project, including methodologies, simulations, and results.

2. **`Best-Of_Pricing_Project.xlsb`**  
   - Excel workbook implementing the simulations and calculations.

3. **`VBA_Code.txt`**  
   - Code for retrieving spot prices, fetching historical data, and running simulations.

---
## Key Features

1. **Data Retrieval**
   - Spot prices for Société Générale and BNP are fetched from **Boursorama** using VBA scripts.
   - Historical price data is downloaded from the **AlphaVantage API**. This data is essential for estimating volatility and correlation between the two assets.

2. **Price Simulation**
   - Monte Carlo simulations are used to generate multiple possible price paths for Société Générale and BNP assets over a given time horizon.
   - The simulation is based on a geometric model that incorporates expected returns, volatility, and random market fluctuations.

3. **Option Pricing**
   - **Best-Of Option Payoff**: This option pays out based on the better-performing asset between Société Générale and BNP, minus a predefined strike price.
   - **Worst-Of Option Payoff**: This option pays out based on the worse-performing asset, relative to the strike price.
   - The prices of both options are calculated by averaging the simulated payoffs and applying a discounting factor for the time value of money.

   **Key Parameters**:
   - **Strike Price**: The predetermined price at which the option becomes profitable.  
   - **Risk-Free Rate**: The theoretical return of an investment with no risk, used for discounting.  
   - **Maturity**: The time until the option expires, typically measured in years.

---

## **Requirements**
- **Software**:
  - Microsoft Excel (with VBA enabled).
  - Access to the AlphaVantage API.
- **Programming Language**:
  - VBA for spot price and historical data retrieval.
- **Data**:
  - Real-time spot prices from Boursorama.
  - Historical data for volatility and correlation calculations.

---

## How to Use

1. **Clone the repository**:
   ```bash
   git clone <repository-url>
   cd Best-Of-Option-Pricing
   ```

2. **Open the Excel file**:
   - Use the file `Best-Of_Pricing_Project.xlsb` to visualize and adjust the parameters.

3. **Run the VBA code**:
   - Enable macros and execute the VBA scripts to retrieve data and run simulations.

4. **Analyze Results**:
   - Check the calculated Best-Of and Worst-Of prices, and compare their behaviors.

---

## Contributors

- **Eloi Martin**  
  - [LinkedIn](https://www.linkedin.com/in/eloi-martin-a20475267/)  
  - [GitHub](https://github.com/EloiMt)

- **Florine Nguyen**  
  - [LinkedIn](https://www.linkedin.com/in/florine-nguyen-325778254/)  
  - [GitHub](https://github.com/FlorineNguyen)


For any questions or suggestions, feel free to open an issue or contact us on LinkedIn/GitHub.

---

## License

This project is licensed under the MIT License. See `LICENSE` for more details.
