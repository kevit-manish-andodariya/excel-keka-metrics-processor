const xlsx = require('xlsx');
const path = require('path');


function readExcel(filePath, sheetName) {
    const workbook = xlsx.readFile(filePath);
    const worksheet = workbook.Sheets[sheetName];
    return xlsx.utils.sheet_to_json(worksheet);
}

function calculateDeliveryQuality(filter) {
    const filePath = path.join(__dirname, 'DC-1.1 - Delivery Quality.xlsx'); // Input Excel file
    const sheetName = 'Sheet1';
    const data = readExcel(filePath, sheetName);

    const filteredData = data.filter((row) => {
        const matchesMonth = filter.month ? row.Month?.toLowerCase() === filter.month.toLowerCase() : true;
        const matchesProject = filter.projectName ? row['Project Name']?.toLowerCase() === filter.projectName.toLowerCase() : true;
        const matchesPerson = filter.person ? row['Delivery owner']?.toLowerCase().includes(filter.person.toLowerCase()) : true;
        return matchesMonth && matchesProject && matchesPerson;
    });

    if (filteredData.length === 0) {
        return `No data found for Month: ${filter.month || 'Any'}, Project: ${filter.projectName || 'Any'}, Person: ${filter.person || 'Any'}.`;
    }

    let totalTestCases = 0;
    let totalPassedTests = 0;

    filteredData.forEach((entry) => {
        totalTestCases += entry['Total Test Cases'] || 0;
        totalPassedTests += entry['Tests Passed'] || 0;
    });

    const percentage = totalTestCases > 0 ? ((totalPassedTests / totalTestCases) * 100).toFixed(2) : 0;

    return `${filter.month || 'Overall'} Delivery Quality Percentage for Project: "${filter.projectName || 'All'}", Person: "${filter.person || 'All'}" is ${percentage}%.`;
}

function calculateOnTimeDelivery(filter) {
    const filePath = path.join(__dirname, 'DC-2.1 - On Time Delivery.xlsx'); // Input Excel file
    const sheetName = 'Data';
    const data = readExcel(filePath, sheetName);

    const filteredData = data.filter((row) => {
        const matchesMonth = filter.month ? row.Month?.toLowerCase() === filter.month.toLowerCase() : true;
        const matchesProject = filter.projectName ? row['Project Name']?.toLowerCase() === filter.projectName.toLowerCase() : true;
        const matchesPerson = filter.person ? row['Delivery Owner']?.toLowerCase().includes(filter.person.toLowerCase()) : true;
        return matchesMonth && matchesProject && matchesPerson;
    });

    if (filteredData.length === 0) {
        return `No data found for Month: ${filter.month || 'Any'}, Project: ${filter.projectName || 'Any'}, Person: ${filter.person || 'Any'}.`;
    }

    let totalDeliveries = 0;
    let onTimeDeliveries = 0;

    filteredData.forEach((entry) => {
        totalDeliveries += 1;

        // Convert the Excel serial date values to JavaScript Date objects
        const scheduledDate = entry['Scheduled Delivery Date'];
        const actualDate = entry['Actual Delivery Date'];

        // Excel serial date starts from 1900-01-01, so we need to adjust accordingly
        const jsScheduledDate = new Date((scheduledDate - 25569) * 86400 * 1000);
        const jsActualDate = new Date((actualDate - 25569) * 86400 * 1000);

        // Check if the actual delivery date is on or before the scheduled date
        const isOnTime = jsActualDate <= jsScheduledDate;

        // If delivery is on time, increment the counter
        if (isOnTime) {
            onTimeDeliveries += 1;
        }
    });

    const percentage = totalDeliveries > 0 ? ((onTimeDeliveries / totalDeliveries) * 100).toFixed(2) : 0;

    return `${filter.month || 'Overall'} On-Time Delivery Percentage for Project: "${filter.projectName || 'All'}", Person: "${filter.person || 'All'}" is ${percentage}%.`;
}

function calculateAverageCodeCoverage(filter) {
    const filePath = path.join(__dirname, 'DC-1.3 - Code Coverage.xlsx'); // Input Excel file
    const sheetName = 'Maintaining Coverage';
    const data = readExcel(filePath, sheetName);

    const filteredData = data.filter((row) => {
        const matchesMonth = filter.month ? row.Month?.toLowerCase() === filter.month.toLowerCase() : true;
        const matchesProject = filter.projectName ? row['Project Name']?.toLowerCase() === filter.projectName.toLowerCase() : true;
        return matchesMonth && matchesProject;
    });

    if (filteredData.length === 0) {
        return `No data found for Month: ${filter.month || 'Any'}, Project: ${filter.projectName || 'Any'}, Repo: ${filter.repo || 'Any'}.`;
    }

    let totalCoverage = 0;
    let count = 0;

    filteredData.forEach((entry) => {
        const coverage = entry['Code Coverage'];
        if (coverage !== 'N/A' && !isNaN(coverage)) {
            totalCoverage += coverage * 100;
            count += 1;
        }
    });

    const averageCoverage = count > 0 ? (totalCoverage / count).toFixed(2) : 0;

    return `${filter.month || 'Overall'} Average Code Coverage for Project: "${filter.projectName || 'All'}", Repo: "${filter.repo || 'All'}" is ${averageCoverage}%.`;
}

function calculateTeamIssueMetrics(filter) {
    const filePath = path.join(__dirname, 'DC-3.1 High Priority Production Issues.xlsx'); // Input Excel file
    const sheetName = 'Data Collection';
    const data = readExcel(filePath, sheetName); // Reading the Excel file

    const parseExcelDate = (value) => {
        // Check if it's already a date object
        if (value instanceof Date) {
            return value;
        }

        // If it's a valid serial date (number), convert it to a JS Date
        if (!isNaN(value)) {
            return new Date((value - 25569) * 86400000); // Excel serial date to JS Date
        }

        // If it's a string that looks like a date, convert it to JS Date
        if (typeof value === 'string' && !isNaN(Date.parse(value))) {
            return new Date(value);
        }

        return null; // Return null if invalid
    };

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

    if (filteredData.length === 0) {
        return `No data found for Month: ${filter.month || 'Any'}, Team: ${filter.team || 'Any'}.`;
    }

    const totalIssues = filteredData.length;
    const onTimeIssues = filteredData.filter((row) => row['On Time Answer( Yes / No)']?.toLowerCase() === 'yes').length;

    let totalAcknowledgmentTime = 0;
    let totalResolutionTime = 0;
    let acknowledgmentCount = 0;
    let resolutionCount = 0;

    filteredData.forEach((row) => {
        const reportedTime = parseExcelDate(row['Reported Time']);
        const acknowledgmentTime = parseExcelDate(row['Initial Acknowledgment Time']);
        const resolutionTime = parseExcelDate(row['Resolution Time']); // Ensure this is parsed correctly

        // Calculate acknowledgment time (in hours)
        if (reportedTime && acknowledgmentTime) {
            const acknowledgmentDiffMillis = acknowledgmentTime - reportedTime;
            const acknowledgmentDiffHours = acknowledgmentDiffMillis / 3600000; // Difference in hours
            totalAcknowledgmentTime += acknowledgmentDiffHours;
            acknowledgmentCount++;
        }

        // Calculate resolution time (in hours)
        if (reportedTime && resolutionTime) {
            const resolutionDiffMillis = resolutionTime - reportedTime;
            const resolutionDiffHours = resolutionDiffMillis / 3600000; // Difference in hours
            totalResolutionTime += resolutionDiffHours;
            resolutionCount++;
        }
    });

    // Compute averages
    const avgAcknowledgmentTimeHours = acknowledgmentCount > 0 ? (totalAcknowledgmentTime / acknowledgmentCount).toFixed(2) : 'N/A';
    const avgAcknowledgmentTimeMinutes = acknowledgmentCount > 0 ? ((totalAcknowledgmentTime / acknowledgmentCount) * 60).toFixed(2) : 'N/A';
    const avgResolutionTime = resolutionCount > 0 ? (totalResolutionTime / resolutionCount).toFixed(2) : 'N/A';
    const avgResolutionTimeMinutes = resolutionCount > 0 ? ((totalResolutionTime / resolutionCount) * 60).toFixed(2) : 'N/A';


    return `
    High priority Production issues Metrics for ${filter.month || 'Overall'} | Team: ${filter.team || 'All'}:
    ------------------------------------------------------------
    - Total Issues: ${totalIssues}
    - On-Time Resolved Issues: ${onTimeIssues}
    - Average Initial Acknowledgment Time: ${avgAcknowledgmentTimeHours} hours (${avgAcknowledgmentTimeMinutes} minutes)
    - Average Resolution Time: ${avgResolutionTime} hours (${avgResolutionTimeMinutes} minutes)
    `;
}


//Inputs
// Get the current month name
function getCurrentMonth() {
    const monthNames = [
        'January', 'February', 'March', 'April', 'May', 'June',
        'July', 'August', 'September', 'October', 'November', 'December'
    ];
    return monthNames[new Date().getMonth()];
}

function createFilter(options = {}) {
    const { month = getCurrentMonth(), projectName = 'Syngenta Planting', person, team } = options;
    return {
        month,
        projectName,
        person,
        team
    };
}

// Centralized function to calculate metrics for a person and month
function calculateMetrics(persons, filterOptions) {
    const defaultFilter = createFilter(filterOptions);

    persons.forEach((person) => {
        const deliveryQuality = calculateDeliveryQuality({ ...defaultFilter, person });
        console.log(`Delivery Quality || `, deliveryQuality);

        const onTimeDelivery = calculateOnTimeDelivery({ ...defaultFilter, person });
        console.log(`On-Time Delivery || `, onTimeDelivery);
    });

    const averageCoverage = calculateAverageCodeCoverage({ ...defaultFilter });
    console.log('Average Code Coverage || ', averageCoverage);

    const prodIssueMetrics = calculateTeamIssueMetrics({ ...defaultFilter, team: filterOptions.projectName });
    console.log('Production Issue Metrics || ', prodIssueMetrics);
}

const persons = ['Manish', 'Yuvraj', 'Gungun'];
const filterOptions = { month: 'November', projectName: 'Syngenta Planting' };

calculateMetrics(persons, filterOptions);
