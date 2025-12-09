 # Excel Data Processor

This Node.js script processes data from multiple Excel files to calculate various software delivery metrics.  It's designed to provide insights into delivery quality, on-time delivery rates, code coverage, and team issue resolution times.

## Requirements

* Node.js and npm (or yarn) installed.
* `xlsx` package:  Install using `npm install xlsx`

## Setup

1. **Clone the repository:** Clone this repository to your local machine.  (e.g., `git clone <your_repo_url>`)
2. **Install dependencies:** Navigate to the project directory and run `npm install`.
3. **Prepare Excel files:** Place the following Excel files (mentioned below) in the same directory as the `index.js` script.  Ensure the sheet names match those hardcoded in the script.
    * `DC-1.1 - Delivery Quality.xlsx`
    * `DC-2.1 - On Time Delivery.xlsx`
    * `DC-1.3 - Code Coverage.xlsx`
    * `DC-3.1 High Priority Production Issues.xlsx`


## Usage

The script calculates metrics based on configurable filters.  You can specify the month, project name, and person (for delivery quality and on-time delivery) or team (for production issue metrics).

**Input Excel Files Structure:**

Make sure your excel files have the following columns. Case sensitivity matters in some areas, please match the column names exactly.

* **`DC-1.1 - Delivery Quality.xlsx`:** `Month`, `Project Name`, `Delivery owner`, `Total Test Cases`, `Tests Passed`
* **`DC-2.1 - On Time Delivery.xlsx`:** `Month`, `Project Name`, `Delivery Owner`, `Scheduled Delivery Date`, `Actual Delivery Date` (Dates should be valid Excel serial dates)
* **`DC-1.3 - Code Coverage.xlsx`:** `Month`, `Project Name`, `Code Coverage` (Numbers or "N/A")
* **`DC-3.1 High Priority Production Issues.xlsx`:** `Reported Time`, `Team Name`, `On Time Answer( Yes / No)`, `Initial Acknowledgment Time`, `Resolution Time` (Dates should be valid Excel serial dates or parsable date strings)


**Running the Script:**

1.  **Direct Execution (with default parameters):** Run `node index.js`. This will use the default parameters: current month, 'Syngenta Planting' project, and a list of predefined persons.

2. **Customizing the Filters:** Modify the `filterOptions` object in the `index.js` file to change the filters.  For example:

```javascript
const persons = ['Manish', 'Yuvraj', 'Gungun']; // Array of people to calculate metrics for
const filterOptions = { 
    month: 'October', // Specify the month 
    projectName: 'Another Project' // Specify the project name
};

calculateMetrics(persons, filterOptions);
```


<div align="center">Made with love from India ❤️</div>


