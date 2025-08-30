/**
 * Enhanced Google Calendar Cleanup Script
 * Advanced productivity tool for managing calendar events with smart filtering
 * and comprehensive reporting capabilities
 */

// Enhanced Configuration with more granular control
const CONFIG = {
  // Time-based settings
  YEARS_TO_KEEP: 1,
  MONTHS_TO_KEEP: null,                // Alternative to years (if set, overrides YEARS_TO_KEEP)
  DAYS_TO_KEEP: null,                  // Alternative for very recent cleanup
  
  // Safety and performance settings
  MAX_DELETES_PER_RUN: 50,
  SLEEP_BETWEEN_DELETES: 1000,
  DRY_RUN: true,                       // Start with dry run for safety
  DELETE_RECURRING_INSTANCES: false,   // Conservative default
  
  // Advanced filtering options
  SKIP_ALL_DAY_EVENTS: false,          // Keep all-day events like holidays
  SKIP_EVENTS_WITH_ATTENDEES: true,    // Keep meeting invites for reference
  SKIP_KEYWORDS: ['birthday', 'anniversary', 'holiday'], // Events to always keep
  ONLY_DELETE_KEYWORDS: [],            // Only delete events containing these keywords
  MIN_TITLE_LENGTH: 1,                 // Skip events with very short titles (likely system-generated)
  
  // Calendar-specific settings
  TARGET_CALENDARS: [],                // Empty = all calendars, or specify calendar names
  EXCLUDE_CALENDARS: ['Birthdays', 'Holidays'], // Calendars to never touch
  
  // Reporting settings
  GENERATE_REPORT: true,
  EMAIL_REPORT: false,                 // Set to true to email results
  REPORT_EMAIL: '',                    // Your email for reports
  
  // Advanced safety features
  CONFIRM_BEFORE_DELETE: false,        // Require manual confirmation (useful for first runs)
  BACKUP_BEFORE_DELETE: false          // Create a backup spreadsheet of deleted events
};

/**
 * Main execution function with enhanced error handling and reporting
 */
function deleteOldCalendarEvents() {
  const startTime = new Date();
  console.log("=== ENHANCED CALENDAR CLEANUP STARTED ===");
  console.log(`Start time: ${startTime.toLocaleString()}`);
  console.log(`Configuration: ${JSON.stringify(CONFIG, null, 2)}`);
  
  let globalResults = {
    totalProcessed: 0,
    totalDeleted: 0,
    totalSkipped: 0,
    totalErrors: 0,
    calendarResults: [],
    startTime: startTime,
    endTime: null,
    errorLog: [],
    backupSheetId: null
  };
  
  try {
    // Validate configuration
    if (!validateConfiguration()) {
      return;
    }
    
    // Get target calendars
    const calendars = getTargetCalendars();
    if (calendars.length === 0) {
      console.log("❌ No calendars available for processing");
      return;
    }
    
    // Calculate date range
    const dateRange = calculateEnhancedDateRange();
    console.log(`Date range: ${dateRange.start.toLocaleDateString()} to ${dateRange.cutoff.toLocaleDateString()}`);
    
    // Create backup if requested
    if (CONFIG.BACKUP_BEFORE_DELETE && !CONFIG.DRY_RUN) {
      globalResults.backupSheetId = createBackupSpreadsheet(calendars, dateRange);
    }
    
    // Process each calendar
    for (const calendar of calendars) {
      const calendarResult = processCalendar(calendar, dateRange);
      globalResults.calendarResults.push(calendarResult);
      globalResults.totalProcessed += calendarResult.processed;
      globalResults.totalDeleted += calendarResult.deleted;
      globalResults.totalSkipped += calendarResult.skipped;
      globalResults.totalErrors += calendarResult.errors;
    }
    
    // Generate comprehensive report
    globalResults.endTime = new Date();
    generateFinalReport(globalResults);
    
    // Email report if requested
    if (CONFIG.EMAIL_REPORT && CONFIG.REPORT_EMAIL) {
      emailReport(globalResults);
    }
    
  } catch (error) {
    console.error(`❌ Fatal error: ${error.message}`);
    console.error(`Stack trace: ${error.stack}`);
    globalResults.errorLog.push(`Fatal error: ${error.message}`);
  }
  
  console.log("=== ENHANCED CALENDAR CLEANUP COMPLETED ===");
  return globalResults;
}

/**
 * Validate configuration settings
 */
function validateConfiguration() {
  console.log("🔍 Validating configuration...");
  
  let isValid = true;
  const warnings = [];
  const errors = [];
  
  // Check date settings
  const dateSettings = [CONFIG.YEARS_TO_KEEP, CONFIG.MONTHS_TO_KEEP, CONFIG.DAYS_TO_KEEP].filter(x => x !== null);
  if (dateSettings.length === 0) {
    errors.push("Must specify at least one time period (YEARS_TO_KEEP, MONTHS_TO_KEEP, or DAYS_TO_KEEP)");
    isValid = false;
  }
  
  // Validate email settings
  if (CONFIG.EMAIL_REPORT && !CONFIG.REPORT_EMAIL) {
    warnings.push("EMAIL_REPORT is true but no REPORT_EMAIL specified");
  }
  
  // Safety warnings
  if (!CONFIG.DRY_RUN && CONFIG.DELETE_RECURRING_INSTANCES) {
    warnings.push("DANGER: You're about to delete recurring event instances in live mode");
  }
  
  if (!CONFIG.DRY_RUN && CONFIG.MAX_DELETES_PER_RUN > 100) {
    warnings.push("Large number of deletions planned - consider smaller batches");
  }
  
  // Log validation results
  if (errors.length > 0) {
    console.error("❌ Configuration errors:");
    errors.forEach(error => console.error(`  - ${error}`));
    return false;
  }
  
  if (warnings.length > 0) {
    console.log("⚠️  Configuration warnings:");
    warnings.forEach(warning => console.log(`  - ${warning}`));
  }
  
  console.log("✅ Configuration validated");
  return isValid;
}

/**
 * Get calendars to process based on configuration
 */
function getTargetCalendars() {
  console.log("📅 Identifying target calendars...");
  
  try {
    const allCalendars = CalendarApp.getAllCalendars();
    let targetCalendars = [];
    
    if (CONFIG.TARGET_CALENDARS.length > 0) {
      // Use only specified calendars
      targetCalendars = allCalendars.filter(cal => 
        CONFIG.TARGET_CALENDARS.includes(cal.getName())
      );
      
      if (targetCalendars.length === 0) {
        console.error("❌ None of the specified TARGET_CALENDARS were found");
        return [];
      }
    } else {
      // Use all calendars except excluded ones
      targetCalendars = allCalendars.filter(cal => 
        !CONFIG.EXCLUDE_CALENDARS.includes(cal.getName())
      );
    }
    
    console.log(`📊 Found ${targetCalendars.length} calendar(s) to process:`);
    targetCalendars.forEach(cal => {
      console.log(`  - "${cal.getName()}"`);
    });
    
    return targetCalendars;
    
  } catch (error) {
    console.error(`❌ Error accessing calendars: ${error.message}`);
    return [];
  }
}

/**
 * Calculate date range with enhanced options
 */
function calculateEnhancedDateRange() {
  const now = new Date();
  const cutoffDate = new Date();
  
  // Use the most specific time setting provided
  if (CONFIG.DAYS_TO_KEEP !== null) {
    cutoffDate.setDate(cutoffDate.getDate() - CONFIG.DAYS_TO_KEEP);
    console.log(`📊 Using DAYS_TO_KEEP: ${CONFIG.DAYS_TO_KEEP} days`);
  } else if (CONFIG.MONTHS_TO_KEEP !== null) {
    cutoffDate.setMonth(cutoffDate.getMonth() - CONFIG.MONTHS_TO_KEEP);
    console.log(`📊 Using MONTHS_TO_KEEP: ${CONFIG.MONTHS_TO_KEEP} months`);
  } else {
    cutoffDate.setFullYear(cutoffDate.getFullYear() - CONFIG.YEARS_TO_KEEP);
    console.log(`📊 Using YEARS_TO_KEEP: ${CONFIG.YEARS_TO_KEEP} years`);
  }
  
  const startDate = new Date(2000, 0, 1); // Far back start date
  
  console.log(`📊 Current date: ${now.toLocaleDateString()}`);
  console.log(`📊 Cutoff date: ${cutoffDate.toLocaleDateString()}`);
  console.log(`📊 Events older than cutoff will be ${CONFIG.DRY_RUN ? 'identified' : 'deleted'}`);
  
  return {
    start: startDate,
    cutoff: cutoffDate,
    now: now
  };
}

/**
 * Process a single calendar
 */
function processCalendar(calendar, dateRange) {
  const calendarName = calendar.getName();
  console.log(`\n🗓️  Processing calendar: "${calendarName}"`);
  
  const result = {
    calendarName: calendarName,
    processed: 0,
    deleted: 0,
    skipped: 0,
    errors: 0,
    errorMessages: [],
    skippedReasons: {},
    sampleEvents: []
  };
  
  try {
    // Get events from this calendar
    const events = calendar.getEvents(dateRange.start, dateRange.cutoff);
    console.log(`📊 Found ${events.length} old events in "${calendarName}"`);
    
    if (events.length === 0) {
      console.log(`✅ No old events in "${calendarName}"`);
      return result;
    }
    
    // Apply smart filtering
    const filteredEvents = applySmartFiltering(events);
    console.log(`📊 After filtering: ${filteredEvents.length} events to process`);
    
    if (filteredEvents.length === 0) {
      console.log(`✅ All events in "${calendarName}" were filtered out (kept for safety)`);
      result.skipped = events.length;
      result.skippedReasons['filtered_out'] = events.length;
      return result;
    }
    
    // Process events
    const maxToProcess = Math.min(filteredEvents.length, CONFIG.MAX_DELETES_PER_RUN);
    console.log(`🗑️  Processing ${maxToProcess} events in "${calendarName}"...`);
    
    for (let i = 0; i < maxToProcess; i++) {
      const eventResult = processIndividualEvent(filteredEvents[i]);
      
      result.processed++;
      if (eventResult.deleted) {
        result.deleted++;
      } else if (eventResult.skipped) {
        result.skipped++;
        const reason = eventResult.reason || 'unknown';
        result.skippedReasons[reason] = (result.skippedReasons[reason] || 0) + 1;
      } else if (eventResult.error) {
        result.errors++;
        result.errorMessages.push(eventResult.errorMessage);
      }
      
      // Store sample events for reporting
      if (result.sampleEvents.length < 5) {
        result.sampleEvents.push({
          title: filteredEvents[i].getTitle() || "(No title)",
          date: filteredEvents[i].getStartTime().toLocaleDateString(),
          action: eventResult.deleted ? 'deleted' : (eventResult.skipped ? 'skipped' : 'error')
        });
      }
      
      // Sleep between operations
      if (i < maxToProcess - 1) {
        Utilities.sleep(CONFIG.SLEEP_BETWEEN_DELETES);
      }
    }
    
  } catch (error) {
    console.error(`❌ Error processing calendar "${calendarName}": ${error.message}`);
    result.errors++;
    result.errorMessages.push(`Calendar processing error: ${error.message}`);
  }
  
  return result;
}

/**
 * Apply smart filtering to events before deletion
 */
function applySmartFiltering(events) {
  console.log("🧠 Applying smart filtering...");
  
  return events.filter(event => {
    try {
      const title = event.getTitle() || "";
      const isAllDay = event.isAllDayEvent();
      const hasAttendees = event.getGuestList().length > 0;
      const isRecurring = event.isRecurringEvent();
      
      // Skip all-day events if configured
      if (CONFIG.SKIP_ALL_DAY_EVENTS && isAllDay) {
        return false;
      }
      
      // Skip events with attendees if configured
      if (CONFIG.SKIP_EVENTS_WITH_ATTENDEES && hasAttendees) {
        return false;
      }
      
      // Skip recurring instances if configured
      if (!CONFIG.DELETE_RECURRING_INSTANCES && isRecurring) {
        return false;
      }
      
      // Skip events with certain keywords
      const titleLower = title.toLowerCase();
      if (CONFIG.SKIP_KEYWORDS.some(keyword => titleLower.includes(keyword.toLowerCase()))) {
        return false;
      }
      
      // Only delete events with specific keywords (if configured)
      if (CONFIG.ONLY_DELETE_KEYWORDS.length > 0) {
        if (!CONFIG.ONLY_DELETE_KEYWORDS.some(keyword => titleLower.includes(keyword.toLowerCase()))) {
          return false;
        }
      }
      
      // Skip very short titles (likely system-generated)
      if (title.length < CONFIG.MIN_TITLE_LENGTH) {
        return false;
      }
      
      return true;
      
    } catch (error) {
      console.log(`⚠️  Error filtering event: ${error.message}`);
      return false; // Skip events we can't analyze
    }
  });
}

/**
 * Process an individual event
 */
function processIndividualEvent(event) {
  const result = {
    deleted: false,
    skipped: false,
    error: false,
    reason: null,
    errorMessage: null
  };
  
  try {
    const title = event.getTitle() || "(No title)";
    const date = event.getStartTime().toLocaleDateString();
    const eventInfo = `"${title}" from ${date}`;
    
    if (CONFIG.DRY_RUN) {
      console.log(`🔍 Would delete: ${eventInfo}`);
      result.deleted = true;
      return result;
    }
    
    // Actual deletion
    event.deleteEvent();
    console.log(`✅ Deleted: ${eventInfo}`);
    result.deleted = true;
    
  } catch (error) {
    if (error.message.toLowerCase().includes('action not allowed')) {
      result.skipped = true;
      result.reason = 'permission_denied';
      console.log(`⏭️  Skipped (no permission): ${event.getTitle()}`);
    } else {
      result.error = true;
      result.errorMessage = error.message;
      console.error(`❌ Error deleting event: ${error.message}`);
    }
  }
  
  return result;
}

/**
 * Create backup spreadsheet of events to be deleted
 */
function createBackupSpreadsheet(calendars, dateRange) {
  console.log("💾 Creating backup spreadsheet...");
  
  try {
    const spreadsheet = SpreadsheetApp.create(`Calendar Backup ${new Date().toISOString().split('T')[0]}`);
    const sheet = spreadsheet.getActiveSheet();
    
    // Set up headers
    sheet.getRange(1, 1, 1, 6).setValues([
      ['Calendar', 'Title', 'Start Date', 'End Date', 'All Day', 'Attendees']
    ]);
    
    let row = 2;
    
    for (const calendar of calendars) {
      const events = calendar.getEvents(dateRange.start, dateRange.cutoff);
      const filteredEvents = applySmartFiltering(events);
      
      for (const event of filteredEvents) {
        sheet.getRange(row, 1, 1, 6).setValues([[
          calendar.getName(),
          event.getTitle() || "(No title)",
          event.getStartTime(),
          event.getEndTime(),
          event.isAllDayEvent(),
          event.getGuestList().map(guest => guest.getEmail()).join(', ')
        ]]);
        row++;
        
        if (row > 1000) break; // Limit backup size
      }
    }
    
    console.log(`✅ Backup created: ${spreadsheet.getUrl()}`);
    return spreadsheet.getId();
    
  } catch (error) {
    console.error(`❌ Error creating backup: ${error.message}`);
    return null;
  }
}

/**
 * Generate comprehensive final report
 */
function generateFinalReport(results) {
  console.log("\n📊 === COMPREHENSIVE RESULTS REPORT ===");
  
  const duration = ((results.endTime - results.startTime) / 1000 / 60).toFixed(2);
  console.log(`Execution time: ${duration} minutes`);
  console.log(`Mode: ${CONFIG.DRY_RUN ? 'DRY RUN (simulation)' : 'LIVE DELETION'}`);
  console.log(`Total events processed: ${results.totalProcessed}`);
  console.log(`Events ${CONFIG.DRY_RUN ? 'identified for deletion' : 'deleted'}: ${results.totalDeleted}`);
  console.log(`Events skipped: ${results.totalSkipped}`);
  console.log(`Errors encountered: ${results.totalErrors}`);
  
  if (results.backupSheetId) {
    const backupUrl = `https://docs.google.com/spreadsheets/d/${results.backupSheetId}`;
    console.log(`Backup spreadsheet: ${backupUrl}`);
  }
  
  // Per-calendar breakdown
  console.log("\n📅 PER-CALENDAR BREAKDOWN:");
  results.calendarResults.forEach(cal => {
    if (cal.processed > 0) {
      console.log(`\n"${cal.calendarName}":`);
      console.log(`  - Processed: ${cal.processed}`);
      console.log(`  - Deleted: ${cal.deleted}`);
      console.log(`  - Skipped: ${cal.skipped}`);
      console.log(`  - Errors: ${cal.errors}`);
      
      if (cal.sampleEvents.length > 0) {
        console.log(`  - Sample events:`);
        cal.sampleEvents.forEach(evt => {
          console.log(`    • "${evt.title}" (${evt.date}) - ${evt.action}`);
        });
      }
    }
  });
  
  // Success rate calculation
  if (results.totalProcessed > 0) {
    const successRate = ((results.totalDeleted / results.totalProcessed) * 100).toFixed(1);
    console.log(`\n📈 Success rate: ${successRate}%`);
  }
  
  // Next steps recommendations
  console.log("\n💡 RECOMMENDATIONS:");
  if (CONFIG.DRY_RUN) {
    console.log("- Review the results above and set DRY_RUN to false when ready");
    console.log("- Consider adjusting filtering options if needed");
  } else if (results.totalDeleted > 0) {
    console.log("- Calendar cleanup completed successfully");
    console.log("- Consider scheduling this script to run monthly for maintenance");
  }
  
  if (results.totalSkipped > 0) {
    console.log("- Some events were skipped (likely due to permissions or filtering)");
    console.log("- Review SKIP_* configuration options if you want to include them");
  }
}

/**
 * Email report functionality
 */
function emailReport(results) {
  try {
    const subject = `Calendar Cleanup Report - ${results.totalDeleted} events ${CONFIG.DRY_RUN ? 'identified' : 'deleted'}`;
    const body = generateEmailBody(results);
    
    GmailApp.sendEmail(CONFIG.REPORT_EMAIL, subject, body);
    console.log(`✅ Report emailed to ${CONFIG.REPORT_EMAIL}`);
    
  } catch (error) {
    console.error(`❌ Error sending email: ${error.message}`);
  }
}

/**
 * Generate email body for report
 */
function generateEmailBody(results) {
  const duration = ((results.endTime - results.startTime) / 1000 / 60).toFixed(2);
  
  let body = `Calendar Cleanup Report\n`;
  body += `========================\n\n`;
  body += `Execution Date: ${results.startTime.toLocaleString()}\n`;
  body += `Duration: ${duration} minutes\n`;
  body += `Mode: ${CONFIG.DRY_RUN ? 'DRY RUN (simulation)' : 'LIVE DELETION'}\n\n`;
  body += `Summary:\n`;
  body += `- Total events processed: ${results.totalProcessed}\n`;
  body += `- Events ${CONFIG.DRY_RUN ? 'identified' : 'deleted'}: ${results.totalDeleted}\n`;
  body += `- Events skipped: ${results.totalSkipped}\n`;
  body += `- Errors: ${results.totalErrors}\n\n`;
  
  if (results.backupSheetId) {
    body += `Backup: https://docs.google.com/spreadsheets/d/${results.backupSheetId}\n\n`;
  }
  
  body += `This is an automated report from your Google Calendar cleanup script.\n`;
  
  return body;
}

// Utility functions for manual testing and analysis

/**
 * Quick test function - run this first
 */
function quickTest() {
  console.log("🧪 Running quick test...");
  
  try {
    const calendar = CalendarApp.getDefaultCalendar();
    console.log(`✅ Default calendar access: "${calendar.getName()}"`);
    
    const today = new Date();
    const events = calendar.getEvents(today, today);
    console.log(`✅ Can read events: ${events.length} events today`);
    
    console.log("✅ Quick test passed! You can now run the main script.");
    
  } catch (error) {
    console.error(`❌ Quick test failed: ${error.message}`);
  }
}

/**
 * Analyze your calendar without deleting anything
 */
function analyzeCalendarOnly() {
  console.log("🔍 Analyzing calendar contents...");
  
  const originalDryRun = CONFIG.DRY_RUN;
  CONFIG.DRY_RUN = true; // Force dry run for analysis
  
  const results = deleteOldCalendarEvents();
  
  CONFIG.DRY_RUN = originalDryRun; // Restore original setting
  
  return results;
}
