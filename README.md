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

## **Key Features**

### **1. Data Retrieval**
- Spot prices are fetched from **Boursorama** using VBA.
- Historical data is downloaded using the **AlphaVantage API** for volatility and correlation estimation.

### **2. Price Simulation**
- **Monte Carlo Simulations**:
  - Simulates price paths for Société Générale and BNP assets using the GBM formula:
    \[
    S_t = S_{t-1} \times \exp\left((\mu - 0.5 \sigma^2) \Delta t + \sigma Z\sqrt{\Delta t}\right)
    \]
    - \(S_t\): Price at time \(t\).
    - \(\mu\): Expected return.
    - \(\sigma\): Volatility.
    - \(Z\): Standard normal random variable.
    - \(\Delta t\): Time increment (\(1/250\)).

### **3. Option Pricing**
- **Best-Of Option Payoff**:  
  \[
  \text{Payoff} = \max\left(\max(S_{\text{SG}}, S_{\text{BNP}}) - K, 0\right)
  \]
- **Worst-Of Option Payoff**:  
  \[
  \text{Payoff} = \max\left(K - \min(S_{\text{SG}}, S_{\text{BNP}}), 0\right)
  \]
- Option prices are calculated by discounting the average payoffs:
  \[
  \text{Price} = e^{-rT} \times \text{Payoff Average}
  \]
  - \(K\): Strike price.
  - \(r\): Risk-free rate.
  - \(T\): Maturity (in years).

### **4. Results**
- Example parameters:
  - Initial price (SG): 26.6
  - Initial price (BNP): 57.4
  - Strike: 30
  - Volatility: 0.2
  - Correlation: -0.2
  - Risk-free rate: 3%
  - Time steps: 250
  - Simulations: 100
- Calculated prices:
  - **Best-Of Call Option**: 31.6
  - **Worst-Of Call Option**: 2.05

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
  - [LinkedIn](https://www.linkedin.com/in/florine-nguyen/)  
  - [GitHub](https://github.com/florine-nguyen)


For any questions or suggestions, feel free to open an issue or contact us on LinkedIn/GitHub.

---

## License

This project is licensed under the MIT License. See `LICENSE` for more details.
