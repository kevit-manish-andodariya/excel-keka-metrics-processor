const xlsx = require('xlsx');
const path = require('path');

function readExcel(filePath, sheetName) {
    try {
        const workbook = xlsx.readFile(filePath);
        const worksheet = workbook.Sheets[sheetName];
        return xlsx.utils.sheet_to_json(worksheet);
    } catch (error) {
        console.error(`Error reading Excel file: ${filePath}, Sheet: ${sheetName}`, error.message);
        return [];
    }
}

function parseExcelDate(value) {
    if (value instanceof Date) return value;
    if (!isNaN(value)) return new Date((value - 25569) * 86400000); // Convert Excel serial date to JS Date
    if (typeof value === 'string' && !isNaN(Date.parse(value))) return new Date(value);
    return null;
}

function noDataMessage(filter, dataFor) {
    return `Month: ${filter.month || 'Overall'} | Person: ${filter.person || 'Any'} | Project: ${filter.projectName || 'Any'} | Data for ${dataFor} : N/A.`;
}

function calculateDeliveryQuality(filter) {
    try {
        const filePath = path.join(__dirname, 'DC-1.1 - Delivery Quality.xlsx');
        const data = readExcel(filePath, 'Sheet1');

        const filteredData = data.filter((row) => {
            const matchesMonth = filter.month ? row.Month?.toLowerCase() === filter.month.toLowerCase() : true;
            const matchesProject = filter.projectName ? row['Project Name']?.toLowerCase() === filter.projectName.toLowerCase() : true;
            const matchesPerson = filter.person ? row['Delivery owner']?.toLowerCase().includes(filter.person.toLowerCase()) : true;
            return matchesMonth && matchesProject && matchesPerson;
        });

        if (!filteredData.length) return noDataMessage(filter, "Delivery Quality");

        const { totalTestCases, totalPassedTests } = filteredData.reduce(
            (acc, entry) => {
                acc.totalTestCases += entry['Total Test Cases'] || 0;
                acc.totalPassedTests += entry['Tests Passed'] || 0;
                return acc;
            },
            { totalTestCases: 0, totalPassedTests: 0 }
        );

        const percentage = totalTestCases > 0 ? ((totalPassedTests / totalTestCases) * 100).toFixed(2) : 0;
        return `Month: ${filter.month || 'Overall'} | Person: ${filter.person || 'Any'} | Project: ${filter.projectName || 'Any'} | Delivery Quality Percentage: ${percentage}%.`;
    } catch (error) {
        console.error('Error in calculateDeliveryQuality:', error.message);
        return 'An error occurred while calculating Delivery Quality.';
    }
}

function calculateOnTimeDelivery(filter) {
    try {
        const filePath = path.join(__dirname, 'DC-2.1 - On Time Delivery.xlsx');
        const data = readExcel(filePath, 'Data');

        const filteredData = data.filter((row) => {
            const matchesMonth = filter.month ? row.Month?.toLowerCase() === filter.month.toLowerCase() : true;
            const matchesProject = filter.projectName ? row['Project Name']?.toLowerCase() === filter.projectName.toLowerCase() : true;
            const matchesPerson = filter.person ? row['Delivery Owner']?.toLowerCase().includes(filter.person.toLowerCase()) : true;
            return matchesMonth && matchesProject && matchesPerson;
        });

        if (!filteredData.length) return noDataMessage(filter, "On-Time Delivery");

        const { totalDeliveries, onTimeDeliveries } = filteredData.reduce(
            (acc, entry) => {
                acc.totalDeliveries++;
                const scheduledDate = parseExcelDate(entry['Scheduled Delivery Date']);
                const actualDate = parseExcelDate(entry['Actual Delivery Date']);
                if (scheduledDate && actualDate && actualDate <= scheduledDate) {
                    acc.onTimeDeliveries++;
                }
                return acc;
            },
            { totalDeliveries: 0, onTimeDeliveries: 0 }
        );

        const percentage = totalDeliveries > 0 ? ((onTimeDeliveries / totalDeliveries) * 100).toFixed(2) : 0;
        return `Month: ${filter.month || 'Overall'} | Person: ${filter.person || 'Any'} | Project: ${filter.projectName || 'Any'} | Delivery Quality Percentage: ${percentage}%.On-Time Delivery Percentage: ${percentage}%.`;
    } catch (error) {
        console.error('Error in calculateOnTimeDelivery:', error.message);
        return 'An error occurred while calculating On-Time Delivery.';
    }
}

function calculateAverageCodeCoverage(filter) {
    try {
        const filePath = path.join(__dirname, 'DC-1.3 - Code Coverage.xlsx');
        const data = readExcel(filePath, 'Maintaining Coverage');

        const filteredData = data.filter((row) => {
            const matchesMonth = filter.month ? row.Month?.toLowerCase() === filter.month.toLowerCase() : true;
            const matchesProject = filter.projectName ? row['Project Name']?.toLowerCase() === filter.projectName.toLowerCase() : true;
            return matchesMonth && matchesProject;
        });

        if (!filteredData.length) return noDataMessage(filter, "Average Code Coverage");

        const { totalCoverage, count } = filteredData.reduce(
            (acc, entry) => {
                const coverage = entry['Code Coverage'];
                if (coverage !== 'N/A' && !isNaN(coverage)) {
                    acc.totalCoverage += coverage * 100;
                    acc.count++;
                }
                return acc;
            },
            { totalCoverage: 0, count: 0 }
        );

        const averageCoverage = count > 0 ? (totalCoverage / count).toFixed(2) : 0;
        return `Month: ${filter.month || 'Overall'} | Person: ${filter.person || 'Any'} | Project: ${filter.projectName || 'Any'} | Average Code Coverage: ${averageCoverage}%.`;
    } catch (error) {
        console.error('Error in calculateAverageCodeCoverage:', error.message);
        return 'An error occurred while calculating Code Coverage.';
    }
}

function calculateTeamIssueMetrics(filter) {
    const filePath = path.join(__dirname, 'DC-3.1 High Priority Production Issues.xlsx'); // Input Excel file
    const sheetName = 'Data Collection';

    try {
        const data = readExcel(filePath, sheetName); // Read Excel file

        // Helper: Parse Excel date
        const parseExcelDate = (value) => {
            try {
                if (value instanceof Date) return value;
                if (!isNaN(value)) return new Date((value - 25569) * 86400000); // Excel serial date to JS Date
                if (typeof value === 'string' && !isNaN(Date.parse(value))) return new Date(value);
            } catch (error) {
                console.warn(`Failed to parse date: ${value}`);
            }
            return null; // Return null if invalid
        };

        // Helper: Filter data based on month and team
        const filteredData = data.filter((row) => {
            const reportedDate = parseExcelDate(row['Reported Time']);
            const matchesMonth = filter.month && reportedDate
                ? reportedDate.toLocaleString('en-US', { month: 'long' }).toLowerCase() === filter.month.toLowerCase()
                : true;
            const matchesTeam = filter.team
                ? row['Team Name']?.toLowerCase() === filter.team.toLowerCase()
                : true;
            return matchesMonth && matchesTeam && reportedDate; // Exclude invalid rows
        });

        if (filteredData.length === 0) return noDataMessage(filter, "High priority Production Issues")

        // Helper: Calculate on-time issues
        const onTimeIssues = filteredData.filter((row) => row['On Time Answer( Yes / No)']?.toLowerCase() === 'yes').length;

        // Helper: Calculate average times
        const calculateAverageTime = (startField, endField) => {
            let totalTime = 0;
            let count = 0;

            filteredData.forEach((row) => {
                const startTime = parseExcelDate(row[startField]);
                const endTime = parseExcelDate(row[endField]);
                if (startTime && endTime) {
                    const timeDiffMillis = endTime - startTime;
                    const timeDiffHours = timeDiffMillis / 3600000; // Difference in hours
                    totalTime += timeDiffHours;
                    count++;
                }
            });

            const averageTimeHours = count > 0 ? (totalTime / count).toFixed(2) : 'N/A';
            const averageTimeMinutes = count > 0 ? ((totalTime / count) * 60).toFixed(2) : 'N/A';

            return { hours: averageTimeHours, minutes: averageTimeMinutes };
        };

        // Metrics calculation
        const totalIssues = filteredData.length;
        const avgAcknowledgmentTime = calculateAverageTime('Reported Time', 'Initial Acknowledgment Time');
        const avgResolutionTime = calculateAverageTime('Reported Time', 'Resolution Time');

        // Return formatted result
        return `Month: ${filter.month || 'Overall'} | Person: ${filter.person || 'Any'} | Project: ${filter.team || 'Any'} | High priority Production Issues Metrics:
        ------------------------------------------------------------
        - Total Issues: ${totalIssues}
        - On-Time Resolved Issues: ${onTimeIssues}
        - Average Initial Acknowledgment Time: ${avgAcknowledgmentTime.hours} hours (${avgAcknowledgmentTime.minutes} minutes)
        - Average Resolution Time: ${avgResolutionTime.hours} hours (${avgResolutionTime.minutes} minutes)
        `;
    } catch (error) {
        console.error(`Error calculating team issue metrics: ${error.message}`);
        return `An error occurred while processing the data. Please check the logs for more details.`;
    }
}

const project = "Syngenta Planting"

function calculateMetrics(persons, filterOptions) {
    const defaultFilter = { month: filterOptions.month };

    persons.forEach((person) => {
        console.log(calculateDeliveryQuality({ ...defaultFilter, person }));
        console.log(calculateOnTimeDelivery({ ...defaultFilter, person }));
    });

    console.log(calculateAverageCodeCoverage({ ...defaultFilter, projectName: project }));
    console.log(calculateTeamIssueMetrics({ ...defaultFilter, team: project }))
}

const persons = ['Manish', 'Yuvraj', 'Gungun'];

// const persons = [
//     'Devansh Kaneriya',
//     'Priya Lakhani',
//     'Arjun Parmar',
//     'Nidhi Kathrotiya',
//     'Ronak Jagani',
//     'Sagar Dhanwani',
//     'Siddh Kothari',
//     'Keval Mehta',
//     'Mayank Parmar',
//     'Vishal Parmar',
//     'Ashish Chandpa',
//     'Riya Sata',
//     'Maurya Valambhiya',
//     'Dhruva Pambhar',
//     'Megha Rana',
//     'Rishit Rajpara',
//     'Nirali Sakdecha',
//     'Riddhi Parmar',
//     'Sagar Nakum'
// ];

// const persons = ["Misri Pandya"]
const filterOptions = { month: 'November' };
calculateMetrics(persons, filterOptions);
