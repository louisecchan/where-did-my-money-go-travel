function extractAgodaBookings() {
  // Get the active spreadsheet and sheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  
  // Set up headers
  const headers = ['Check-in Date', 'Check-out Date', 'Length of Stay', 'Total Charge', 'Price per Night'];
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  const currentHeaders = headerRange.getValues()[0];
  
  // Check if headers need to be set
  if (currentHeaders.join('') === '' || !currentHeaders.every((header, index) => header === headers[index])) {
    headerRange.setValues([headers]);
    // Format header row
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#f3f3f3');
  }
  
  // Search for Agoda confirmation emails
  const searchQuery = 'from:agoda.com subject:"Booking confirmation"';
  const threads = GmailApp.search(searchQuery, 0, 100);
  
  // Process each email thread
  threads.forEach(thread => {
    const messages = thread.getMessages();
    messages.forEach(message => {
      const subject = message.getSubject();
      const body = message.getPlainBody();
      
      // Extract information using regex patterns
      const checkInDate = extractCheckInDate(body);
      const checkOutDate = extractCheckOutDate(body);
      const totalCharge = extractTotalCharge(body);
      
      // Calculate length of stay if both dates are available
      let lengthOfStay = '';
      let pricePerNight = '';
      if (checkInDate && checkOutDate) {
        lengthOfStay = calculateLengthOfStay(checkInDate, checkOutDate);
        // Calculate price per night if we have total charge and length of stay
        if (totalCharge) {
          const convertedTotal = convertToHKD(totalCharge);
          pricePerNight = calculatePricePerNight(convertedTotal, lengthOfStay);
        }
      }
      
      // Add data to sheet if all information is found
      if (checkInDate && checkOutDate && totalCharge) {
        const lastRow = sheet.getLastRow();
        const convertedTotal = convertToHKD(totalCharge);
        sheet.getRange(lastRow + 1, 1, 1, 5).setValues([[checkInDate, checkOutDate, lengthOfStay, convertedTotal, pricePerNight]]);
      }
    });
  });
}

function convertToHKD(amountStr) {
  try {
    // Extract currency and amount - handle various formats like "IDR 1,234,567" or "1,234,567 IDR"
    const currencyMatch = amountStr.match(/([A-Z]{3})\s*([0-9,.]+)|([0-9,.]+)\s*([A-Z]{3})/);
    if (!currencyMatch) return amountStr; // Return original if no currency found
    
    // Determine which group contains the currency and amount
    const currency = currencyMatch[1] || currencyMatch[4];
    const amount = parseFloat((currencyMatch[2] || currencyMatch[3]).replace(/,/g, ''));
    
    if (isNaN(amount)) return amountStr;
    
    // If already HKD, return formatted amount
    if (currency === 'HKD') {
      return 'HKD ' + amount.toFixed(2);
    }
    
    // Get exchange rate from Google Finance
    const exchangeRate = getExchangeRate(currency, 'HKD');
    if (!exchangeRate) {
      Logger.log(`Failed to get exchange rate for ${currency} to HKD`);
      return amountStr; // Return original if conversion fails
    }
    
    // Convert to HKD
    let hkdAmount;
    if (currency === 'IDR') {
      // For IDR, divide by 2083.33 to get HKD (1 HKD ≈ 2083.33 IDR)
      hkdAmount = amount / 2083.33;
    } else {
      // For other currencies, use the exchange rate
      hkdAmount = amount * exchangeRate;
    }
    
    Logger.log(`Converted ${amount} ${currency} to ${hkdAmount} HKD`);
    return 'HKD ' + hkdAmount.toFixed(2);
  } catch (e) {
    Logger.log(`Error in convertToHKD: ${e.toString()}`);
    return amountStr; // Return original if any error occurs
  }
}

function getExchangeRate(fromCurrency, toCurrency) {
  try {
    // For IDR to HKD, we need to get the rate in the correct direction
    if (fromCurrency === 'IDR' && toCurrency === 'HKD') {
      // Use a fixed rate for IDR to HKD (1 HKD = 2083.33 IDR)
      return 2083.33;
    }
    
    // Use Google Finance to get exchange rate
    const url = `https://www.google.com/finance/quote/${fromCurrency}-${toCurrency}`;
    const response = UrlFetchApp.fetch(url);
    const content = response.getContentText();
    
    // Extract exchange rate from response - look for the actual rate value
    const rateMatch = content.match(/data-last-price="([0-9.]+)"/);
    if (rateMatch) {
      const rate = parseFloat(rateMatch[1]);
      Logger.log(`Found exchange rate for ${fromCurrency}-${toCurrency}: ${rate}`);
      return rate;
    }
    
    // Alternative pattern if the first one fails
    const altRateMatch = content.match(/data-price="([0-9.]+)"/);
    if (altRateMatch) {
      const rate = parseFloat(altRateMatch[1]);
      Logger.log(`Found exchange rate (alt) for ${fromCurrency}-${toCurrency}: ${rate}`);
      return rate;
    }
    
    Logger.log(`No exchange rate found for ${fromCurrency}-${toCurrency}`);
    return null;
  } catch (e) {
    Logger.log(`Error in getExchangeRate: ${e.toString()}`);
    return null;
  }
}

function calculateLengthOfStay(checkInDate, checkOutDate) {
  try {
    // Parse the dates
    const checkIn = new Date(checkInDate);
    const checkOut = new Date(checkOutDate);
    
    // Calculate the difference in milliseconds
    const diffTime = Math.abs(checkOut - checkIn);
    // Convert to days
    const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
    
    return diffDays + ' night' + (diffDays !== 1 ? 's' : '');
  } catch (e) {
    return '';
  }
}

function calculatePricePerNight(totalCharge, lengthOfStay) {
  try {
    // Extract the number of nights from lengthOfStay string
    const nights = parseInt(lengthOfStay);
    if (isNaN(nights) || nights <= 0) return '';
    
    // Extract the numeric value from total charge
    const totalAmount = parseFloat(totalCharge.replace(/[^0-9.]/g, ''));
    if (isNaN(totalAmount)) return '';
    
    // Calculate price per night
    const pricePerNight = totalAmount / nights;
    
    // Format the result with 2 decimal places and HKD symbol
    return 'HKD ' + pricePerNight.toFixed(2);
  } catch (e) {
    return '';
  }
}

function extractCheckInDate(body) {
  // Look for check-in date in various formats
  const patterns = [
    /id="checkin-date">\s*<p[^>]*>\s*([^<]+)/i,
    /Check in:?\s*([^\n]+)/i,
    /Arrival:?\s*([^\n]+)/i
  ];
  
  for (const pattern of patterns) {
    const match = body.match(pattern);
    if (match) {
      const dateStr = match[1].trim();
      // Remove any time information in parentheses
      return dateStr.replace(/\s*\([^)]*\)/, '').trim();
    }
  }
  return null;
}

function extractCheckOutDate(body) {
  // Look for check-out date in various formats
  const patterns = [
    /id="checkout-date">\s*<p[^>]*>\s*([^<]+)/i,
    /Check out:?\s*([^\n]+)/i,
    /Departure:?\s*([^\n]+)/i
  ];
  
  for (const pattern of patterns) {
    const match = body.match(pattern);
    if (match) {
      const dateStr = match[1].trim();
      // Remove any time information in parentheses
      return dateStr.replace(/\s*\([^)]*\)/, '').trim();
    }
  }
  return null;
}

function extractTotalCharge(body) {
  // Look for total charge in various formats
  const patterns = [
    /Total Charge:?\s*([^\n]+)/i,
    /Total Amount:?\s*([^\n]+)/i,
    /Total Payment:?\s*([^\n]+)/i,
    /Total:?\s*([^\n]+)/i
  ];
  
  for (const pattern of patterns) {
    const match = body.match(pattern);
    if (match) {
      return match[1].trim();
    }
  }
  return null;
}

// Create a menu item to run the script
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Agoda Bookings')
    .addItem('Extract Bookings', 'extractAgodaBookings')
    .addItem('Schedule Weekly Run', 'createWeeklyTrigger')
    .addItem('Stop Weekly Run', 'deleteTriggers')
    .addToUi();
}

// Create a weekly trigger
function createWeeklyTrigger() {
  // Delete any existing triggers first
  deleteTriggers();
  
  // Create a new weekly trigger
  ScriptApp.newTrigger('extractAgodaBookings')
    .timeBased()
    .everyWeeks(1)
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(9) // Run at 9 AM
    .create();
  
  // Show confirmation
  const ui = SpreadsheetApp.getUi();
  ui.alert('Weekly Schedule Set', 'The script will now run every Monday at 9 AM.', ui.ButtonSet.OK);
}

// Delete all triggers
function deleteTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    ScriptApp.deleteTrigger(trigger);
  });
  
  // Show confirmation
  const ui = SpreadsheetApp.getUi();
  ui.alert('Schedule Cleared', 'All scheduled runs have been cancelled.', ui.ButtonSet.OK);
}
