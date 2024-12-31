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

// Helper function to normalize strings
const normalizeString = (str) => {
    return str
        ? str
            .toLowerCase()               // Convert to lowercase
            .replace(/\s+/g, ' ')         // Replace multiple spaces with a single space
            .trim()                       // Remove leading/trailing spaces
            .replace(/[^\w\s]/g, '')      // Remove special characters (optional, can modify)
        : '';
};

function parseExcelDate(value) {
    if (value instanceof Date) return value;
    if (!isNaN(value)) return new Date((value - 25569) * 86400000); // Convert Excel serial date to JS Date
    if (typeof value === 'string' && !isNaN(Date.parse(value))) return new Date(value);
    return null;
}

function formatExcelDate(excelDate) {
    const date = new Date((excelDate - 25569) * 86400000);
    return date.toLocaleString();
}

function noDataMessage(filter, dataFor) {
    return `Month: ${filter.month || 'Overall'} | Person: ${filter.person || 'Any'} | Project: ${filter.projectName || 'Any'} | Data for ${dataFor} : N/A.`;
}

const quarterMonths = {
    'Q1': ['january', 'february', 'march'],
    'Q2': ['april', 'may', 'june'],
    'Q3': ['july', 'august', 'september'],
    'Q4': ['october', 'november', 'december']
};

function calculateDeliveryQuality(filter) {
    try {
        const filePath = path.join(__dirname, 'DC-1.1 - Delivery Quality.xlsx');
        const data = readExcel(filePath, 'Data');

        const filteredData = data.filter((row) => {
            const month = row.Month?.toLowerCase();
            const matchesQuarter = filter.quarter ? quarterMonths[filter.quarter].includes(month) : true;
            const matchesProject = filter.projectName ? normalizeString(row['Project Name']) === normalizeString(filter.projectName) : true;
            const matchesPerson = filter.person ? normalizeString(row['Delivery owner']).includes(normalizeString(filter.person)) : false;
            return matchesQuarter && matchesProject && matchesPerson;
        });

        if (!filteredData.length) return noDataMessage(filter, "Delivery Quality");

        console.log(`Entries considered: ${filteredData.length} for ${filter.quarter} | ${filter.person}`)

        const {totalTestCases, totalPassedTests} = filteredData.reduce(
            (acc, entry) => {
                acc.totalTestCases += entry['Total Test Cases'] || 0;
                acc.totalPassedTests += entry['Tests Passed'] || 0;
                return acc;
            },
            {totalTestCases: 0, totalPassedTests: 0}
        );

        const percentage = totalTestCases > 0 ? ((totalPassedTests / totalTestCases) * 100).toFixed(2) : 0;
        return {
            count: filteredData.length,
            Quarter: filter.quarter,
            filteredData,
            metrics: `${percentage}%`,
            message: `Quarter: ${filter.quarter || 'Overall'} | Person: ${filter.person || 'Any'} | Project: ${filter.projectName || 'Any'} | Delivery Quality Percentage: ${percentage}%.`
        }
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
            const month = row.Month?.toLowerCase();
            const matchesQuarter = filter.quarter ? quarterMonths[filter.quarter].includes(month) : true;
            const matchesProject = filter.projectName ? row['Project Name']?.toLowerCase() === filter.projectName.toLowerCase() : true;
            const matchesPerson = filter.person ? row['Delivery Owner']?.toLowerCase().includes(filter.person.toLowerCase()) : false;
            return matchesQuarter && matchesProject && matchesPerson;
        });

        if (!filteredData.length) return noDataMessage(filter, "On-Time Delivery");

        console.log(`Entries considered: ${filteredData.length} for ${filter.quarter} | ${filter.person}`)

        const {totalDeliveries, onTimeDeliveries} = filteredData.reduce(
            (acc, entry) => {
                acc.totalDeliveries++;
                const scheduledDate = parseExcelDate(entry['Scheduled Delivery Date']);
                const actualDate = parseExcelDate(entry['Actual Delivery Date']);
                if (scheduledDate && actualDate && (actualDate <= scheduledDate)) {
                    acc.onTimeDeliveries++;
                }
                return acc;
            },
            {totalDeliveries: 0, onTimeDeliveries: 0}
        );

        const percentage = totalDeliveries > 0 ? ((onTimeDeliveries / totalDeliveries) * 100).toFixed(2) : 0;
        return {
            count: filteredData.length,
            Quarter: filter.quarter,
            filteredData,
            metrics: `${percentage}%`,
            message: `Quarter: ${filter.quarter || 'Overall'} | Person: ${filter.person || 'Any'} | Project: ${filter.projectName || 'Any'} | On-Time Delivery Percentage: ${percentage}%.`
        };
    } catch (error) {
        console.error('Error in calculateOnTimeDelivery:', error.message);
        return 'An error occurred while calculating On-Time Delivery.';
    }
}

function calculateAverageCodeCoverage(filter) {
    try {
        const filePath = path.join(__dirname, 'DC-1.3 - Code Coverage.xlsx');
        const data = readExcel(filePath, 'Data');

        const filteredData = data.filter((row) => {
            const month = row.Month?.toLowerCase();
            const matchesQuarter = filter.quarter ? quarterMonths[filter.quarter].includes(month) : true;
            const matchesProject = filter.projectName ? row['Project Name']?.toLowerCase() === filter.projectName.toLowerCase() : true;
            return matchesQuarter && matchesProject;
        });

        if (!filteredData.length) return noDataMessage(filter, "Average Code Coverage");

        console.log(`Entries considered: ${filteredData.length} for ${filter.quarter} | ${filter.projectName}`)

        const {totalCoverage, count} = filteredData.reduce(
            (acc, entry) => {
                const coverage = entry['Code Coverage'];
                if (coverage !== 'N/A' && !isNaN(coverage)) {
                    acc.totalCoverage += coverage * 100;
                    acc.count++;
                }
                return acc;
            },
            {totalCoverage: 0, count: 0}
        );

        const averageCoverage = count > 0 ? (totalCoverage / count).toFixed(2) : 0;
        return {
            count: filteredData.length,
            Quarter: filter.quarter,
            filteredData,
            metrics: `${averageCoverage}%`,
            message: `Quarter: ${filter.quarter || 'Overall'} | Person: ${filter.person || 'Any'} | Project: ${filter.projectName || 'Any'} | Average Code Coverage: ${averageCoverage}%.`
        };
    } catch (error) {
        console.error('Error in calculateAverageCodeCoverage:', error.message);
        return 'An error occurred while calculating Code Coverage.';
    }
}

function calculateTeamIssueMetrics(filter) {
    const filePath = path.join(__dirname, 'DC-3.1 High Priority Production Issues.xlsx'); // Input Excel file
    const sheetName = 'Data';

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
        const dataToLog = []

        // Helper: Filter data based on quarter and team
        const filteredData = data.filter((row) => {
            const reportedDate = parseExcelDate(row['Reported Time']);
            const month = reportedDate ? reportedDate.toLocaleString('en-US', {month: 'long'}).toLowerCase() : null;
            const matchesQuarter = filter.quarter ? quarterMonths[filter.quarter].includes(month) : true;
            const matchesTeam = filter.team ? row['Team Name']?.toLowerCase() === filter.team.toLowerCase() : true;
            if (matchesTeam && !(matchesQuarter && matchesTeam && reportedDate)) {
                dataToLog.push(row)
            }
            return matchesQuarter && matchesTeam && reportedDate; // Exclude invalid rows
        });

        console.log({dataToLog, filteredData})

        if (filteredData.length === 0) return noDataMessage(filter, "High priority Production Issues");

        console.log(`Entries considered: ${filteredData.length} for ${filter.quarter} | ${filter.team}`)

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

            return {hours: averageTimeHours, minutes: averageTimeMinutes};
        };

        // Metrics calculation
        const totalIssues = filteredData.length;
        const avgAcknowledgmentTime = calculateAverageTime('Reported Time', 'Initial Acknowledgment Time');
        const avgResolutionTime = calculateAverageTime('Reported Time', 'Resolution Time');

        // Return formatted result
        return {
            count: filteredData.length,
            Quarter: filter.quarter,
            filteredData,
            metrics: `${avgAcknowledgmentTime.hours} hours (${avgAcknowledgmentTime.minutes} minutes)`,
            message: `Quarter: ${filter.quarter || 'Overall'} | Person: ${filter.person || 'Any'} | Project: ${filter.team || 'Any'} | High priority Production Issues Metrics:
        ------------------------------------------------------------
        - Total Issues: ${totalIssues}
        - On-Time Resolved Issues: ${onTimeIssues}
        - Average Initial Acknowledgment Time: ${avgAcknowledgmentTime.hours} hours (${avgAcknowledgmentTime.minutes} minutes)
        - Average Resolution Time: ${avgResolutionTime.hours} hours (${avgResolutionTime.minutes} minutes)
        `
        }
    } catch (error) {
        console.error(`Error calculating team issue metrics: ${error.message}`);
        return `An error occurred while processing the data. Please check the logs for more details.`;
    }
}

const project = ["Syngenta Planting"];

const fs = require('fs');

function writeResultsToFile(results, fileName) {
    try {
        fs.writeFileSync(fileName, JSON.stringify(results, null, 2));
        console.log(`Results written to ${fileName}`);
    } catch (error) {
        console.error(`Error writing results to file: ${fileName}`, error.message);
    }
}

function calculateMetrics(persons, filterOptions) {
    const defaultFilter = {quarter: filterOptions.quarter};
    const results = [];

    persons.forEach((person) => {
        const deliveryQuality = calculateDeliveryQuality({...defaultFilter, person});
        const onTimeDelivery = calculateOnTimeDelivery({...defaultFilter, person});
        results.push({
            person,
            deliveryQuality,
            onTimeDelivery
        });
    });

    project.forEach(e => {
        const averageCodeCoverage = calculateAverageCodeCoverage({...defaultFilter, projectName: e});
        const teamIssueMetrics = calculateTeamIssueMetrics({...defaultFilter, team: e});
        results.push({
            project: e,
            averageCodeCoverage,
            highPriorityProductionIssues: teamIssueMetrics
        });
    });
    return results
}

const persons = ["Gungun Udhani", "Yuvraj Kanakiya", "Manish Andodariya"];
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
// const persons = ["Manish Andodariya", "Misri Pandya", "Siddharth Kanjaria", "Siddharth Singh"];
const filterOptions = {quarter: 'Q4'};
const results = calculateMetrics(persons, filterOptions);
writeResultsToFile(results, path.join(__dirname, 'results.json'));
