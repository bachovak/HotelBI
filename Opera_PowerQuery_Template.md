# Opera PMS — Power Query Cleaning Template
# Hotel BI by [Your Name]
# Version 1.1 — Updated from hotelbimodel.bim
#
# HOW TO USE THIS FILE
# Each section below is a standalone Power Query function or query.
# Copy the M code into the Advanced Editor in Power BI Desktop.
# Start with the Parameters, then the Helper Functions, then the main queries.
# ─────────────────────────────────────────────────────────────────────────────


═══════════════════════════════════════════════════════════════
SECTION 1 — PARAMETERS
(Create each of these via Home > Manage Parameters in Power BI)
═══════════════════════════════════════════════════════════════

Parameter Name        | Type   | Default Value         | Description
──────────────────────|────────|───────────────────────|────────────────────────────────────────
HotelName             | Text   | "Hotel Name"          | Used in report titles and page headers
TotalRooms            | Number | 80                    | Total available rooms in the property
FiscalYearStartMonth  | Number | 1                     | 1=Jan, 4=Apr etc. Adjust per hotel
CurrencySymbol        | Text   | "€"                   | For display formatting
DateFormat            | Text   | "DMY"                 | "DMY" or "MDY" — check Opera locale
FolderPath_Flash      | Text   | "C:\HotelBI\Daily Flash"  | Folder where daily flash exports are dropped
FolderPath_Reservations | Text | "C:\HotelBI\Reservations" | Folder for reservation exports
FolderPath_Segments   | Text   | "C:\HotelBI\Segments"     | Folder for segment/source exports


═══════════════════════════════════════════════════════════════
SECTION 2 — HELPER FUNCTIONS
(Create each as a Blank Query. Name them exactly as shown.)
═══════════════════════════════════════════════════════════════

──────────────────────────────────────────
FUNCTION: fnParseOperaDate
Purpose: Handles Opera's date format inconsistencies.
         Opera can export dates as text in DD/MM/YYYY, MM/DD/YYYY, or DD-MMM-YYYY.
         This function tries all three before erroring.
──────────────────────────────────────────

let
    fnParseOperaDate = (dateText as text, format as text) as date =>
        let
            // Remove any surrounding whitespace
            cleaned = Text.Trim(dateText),

            // Split on common separators
            parts = 
                if Text.Contains(cleaned, "/") then Text.Split(cleaned, "/")
                else if Text.Contains(cleaned, "-") then Text.Split(cleaned, "-")
                else error "Unrecognised date separator in: " & cleaned,

            // Parse based on format parameter (DMY or MDY)
            parsed =
                if format = "DMY" then
                    // DD/MM/YYYY
                    #date(
                        Number.From(parts{2}),
                        Number.From(parts{1}),
                        Number.From(parts{0})
                    )
                else if format = "MDY" then
                    // MM/DD/YYYY
                    #date(
                        Number.From(parts{2}),
                        Number.From(parts{0}),
                        Number.From(parts{1})
                    )
                else
                    error "Invalid format parameter. Use DMY or MDY."
        in
            parsed
in
    fnParseOperaDate


──────────────────────────────────────────
FUNCTION: fnCleanNumeric
Purpose: Strips currency symbols, % signs, commas and whitespace from
         Opera's numeric columns which are often exported as text.
         Returns a clean decimal number or null if unparseable.
──────────────────────────────────────────

let
    fnCleanNumeric = (rawValue as any) as any =>
        let
            // Return null for genuinely null/empty inputs
            result = 
                if rawValue = null then null
                else
                    let
                        asText = Text.Trim(Text.From(rawValue)),
                        
                        // Strip common non-numeric characters Opera includes
                        stripped = 
                            Text.Replace(
                            Text.Replace(
                            Text.Replace(
                            Text.Replace(
                            Text.Replace(
                                asText,
                            "€", ""),
                            "$", ""),
                            "%", ""),
                            ",", ""),
                            " ", ""),
                        
                        // Handle bracketed negatives like (1234) used in some Opera reports
                        isNegative = Text.StartsWith(stripped, "(") and Text.EndsWith(stripped, ")"),
                        
                        withoutBrackets = 
                            if isNegative 
                            then Text.Middle(stripped, 1, Text.Length(stripped) - 2)
                            else stripped,
                        
                        asNumber = Number.FromText(withoutBrackets),
                        
                        finalValue = if isNegative then asNumber * -1 else asNumber
                    in
                        try finalValue otherwise null
        in
            result
in
    fnCleanNumeric


──────────────────────────────────────────
FUNCTION: fnRemoveOperaJunk
Purpose: Removes the header rows, footer rows, totals rows, and blank rows
         that Opera inserts around actual data in its Excel exports.
         Pass in a table and the column name that should contain dates or
         numeric IDs — rows where this is null/blank or contains known
         Opera junk strings will be removed.
──────────────────────────────────────────

let
    fnRemoveOperaJunk = (sourceTable as table, keyColumn as text) as table =>
        let
            // List of strings Opera commonly puts in rows that aren't data
            junkStrings = {
                "Total", "Grand Total", "Totals", "Sub Total", "Sub-Total",
                "Report:", "Property:", "Date:", "Page:", "Opera",
                "Run Date", "Run Time", "Period:", "Report Name",
                "*** End of Report ***", "---"
            },

            // Step 1: Remove rows where key column is null
            removeNulls = Table.SelectRows(
                sourceTable, 
                each Record.Field(_, keyColumn) <> null 
                    and Record.Field(_, keyColumn) <> ""
            ),

            // Step 2: Remove rows where key column contains known junk strings
            removeJunk = Table.SelectRows(
                removeNulls,
                each not List.AnyTrue(
                    List.Transform(
                        junkStrings,
                        (junk) => Text.Contains(
                            Text.From(Record.Field(_, keyColumn)), 
                            junk, 
                            Comparer.OrdinalIgnoreCase
                        )
                    )
                )
            ),

            // Step 3: Remove completely empty rows
            removeBlanks = Table.SelectRows(
                removeJunk,
                each List.AnyTrue(
                    List.Transform(
                        Record.FieldValues(_),
                        (v) => v <> null and v <> ""
                    )
                )
            )
        in
            removeBlanks
in
    fnRemoveOperaJunk


──────────────────────────────────────────
FUNCTION: fnGetFileDate
Purpose: Extracts the date from an Opera export filename.
         Opera typically names files like "DailyFlash_20240315.xlsx"
         or "Flash_15-03-2024.xlsx". This tries to extract a date
         from the filename, falling back to file modified date.
──────────────────────────────────────────

let
    fnGetFileDate = (fileName as text, fileModifiedDate as datetime) as date =>
        let
            // Try to find an 8-digit sequence that looks like YYYYMMDD
            digits = Text.Select(fileName, {"0".."9"}),
            
            hasEightDigits = Text.Length(digits) >= 8,
            
            extractedDate =
                if hasEightDigits then
                    let
                        // Take first 8 digits and try YYYYMMDD format
                        first8 = Text.Start(digits, 8),
                        yr = Number.From(Text.Start(first8, 4)),
                        mo = Number.From(Text.Middle(first8, 4, 2)),
                        dy = Number.From(Text.Middle(first8, 6, 2)),
                        isValidDate = yr >= 2000 and yr <= 2100 and mo >= 1 and mo <= 12 and dy >= 1 and dy <= 31
                    in
                        if isValidDate then #date(yr, mo, dy) else Date.From(fileModifiedDate)
                else
                    Date.From(fileModifiedDate)
        in
            extractedDate
in
    fnGetFileDate


═══════════════════════════════════════════════════════════════
SECTION 3 — MAIN QUERIES
═══════════════════════════════════════════════════════════════

──────────────────────────────────────────
QUERY: Fact_DailyFlash
Purpose: Loads and cleans all Daily Flash / Manager's Report exports
         from the FolderPath_Flash folder.

IMPORTANT — Before using this query:
  1. Drop one sample Flash export into your folder
  2. Run the query once to see the raw column structure
  3. Update the column rename step below to match YOUR Opera's exact column names
  4. The column names below are representative — Opera column names vary by property config
──────────────────────────────────────────

let
    // ── Step 1: Load all Excel files from the Flash folder ──
    Source = Folder.Files(FolderPath_Flash),

    ExcelFilesOnly = Table.SelectRows(
        Source,
        each Text.Lower(Text.End([Name], 5)) = ".xlsx"
           or Text.Lower(Text.End([Name], 4)) = ".xls"
    ),

    // ── Step 2: Add file date from filename ──
    AddFileDate = Table.AddColumn(
        ExcelFilesOnly,
        "FileDate",
        each fnGetFileDate([Name], [Date modified]),
        type date
    ),

    // ── Step 3: Define per-file processing function ──
    // FileDate is NOT included here — it already exists on the outer table
    ProcessFile = (fileContent as binary) as table =>
        let
            Workbook = Excel.Workbook(fileContent, true, true),

            // Get the Report sheet
            Sheet =
                let
                    matching = Table.SelectRows(Workbook, each [Name] = "Report")
                in
                    if Table.RowCount(matching) > 0
                    then matching{0}[Data]
                    else Workbook{0}[Data],

            // All rows as a plain list for easy row-by-row access
            AllRows = Table.ToRows(Sheet),

            // Helper: find a row by metric name and return its Actual value (column B)
            FindRow = (metricName as text) as any =>
                let
                    matched = List.Select(AllRows, each _{0} = metricName)
                in
                    if List.Count(matched) > 0 then matched{0}{1} else null,

            // Parse date from Business Date row
            RawDate = FindRow("Business Date"),
            ParsedDate =
                if RawDate = null then null
                else if Value.Is(RawDate, type date) then RawDate
                else if Value.Is(RawDate, type number) then Date.From(RawDate)
                else try fnParseOperaDate(Text.From(RawDate), DateFormat) otherwise null,

            // Build one-row result table directly
            Result = #table(
                {"Date",
                 "RoomsAvailable", "RoomsOccupied", "RoomsOOO",
                 "RoomsComp", "RoomsHouseUse",
                 "OccupancyPct_Opera", "ADR_Opera", "RevPAR_Opera",
                 "RoomRevenue", "FBRevenue", "OtherRevenue", "TotalRevenue"},
                {{
                    ParsedDate,
                    FindRow("Rooms Available"),
                    FindRow("Rooms Occupied"),
                    FindRow("Rooms Out of Order"),
                    FindRow("Complimentary Rooms"),
                    FindRow("House Use Rooms"),
                    FindRow("Occupancy %"),
                    FindRow("ADR"),
                    FindRow("RevPAR"),
                    FindRow("Room Revenue"),
                    FindRow("F&B Revenue"),
                    FindRow("Other Revenue"),
                    FindRow("TOTAL REVENUE")
                }}
            )
        in
            Result,

    // ── Step 4: Apply function to every file ──
    AddProcessed = Table.AddColumn(
        AddFileDate,
        "Processed",
        each ProcessFile([Content])
    ),

    // ── Step 5: Expand results — FileDate comes from outer table automatically ──
    Expanded = Table.ExpandTableColumn(
        AddProcessed,
        "Processed",
        {"Date",
         "RoomsAvailable", "RoomsOccupied", "RoomsOOO",
         "RoomsComp", "RoomsHouseUse",
         "OccupancyPct_Opera", "ADR_Opera", "RevPAR_Opera",
         "RoomRevenue", "FBRevenue", "OtherRevenue", "TotalRevenue"},
        {"Date",
         "RoomsAvailable", "RoomsOccupied", "RoomsOOO",
         "RoomsComp", "RoomsHouseUse",
         "OccupancyPct_Opera", "ADR_Opera", "RevPAR_Opera",
         "RoomRevenue", "FBRevenue", "OtherRevenue", "TotalRevenue"}
    ),

    // ── Step 6: Set column types ──
    SetTypes = Table.TransformColumnTypes(
        Expanded,
        {
            {"Date", type date},
            {"FileDate", type date},
            {"RoomsAvailable", type number},
            {"RoomsOccupied", type number},
            {"RoomsOOO", type number},
            {"RoomsComp", type number},
            {"RoomsHouseUse", type number},
            {"OccupancyPct_Opera", type number},
            {"ADR_Opera", type number},
            {"RevPAR_Opera", type number},
            {"RoomRevenue", type number},
            {"FBRevenue", type number},
            {"OtherRevenue", type number},
            {"TotalRevenue", type number}
        }
    ),

    // ── Step 7: Recalculate KPIs from source numbers ──
    AddOccupancy = Table.AddColumn(
        SetTypes,
        "OccupancyPct",
        each
            let
                avail = if [RoomsAvailable] = null or [RoomsAvailable] = 0
                        then TotalRooms else [RoomsAvailable],
                occ   = if [RoomsOccupied] = null then 0 else [RoomsOccupied]
            in
                if avail = 0 then null else occ / avail,
        type number
    ),

    AddADR = Table.AddColumn(
        AddOccupancy,
        "ADR",
        each
            let
                rev = if [RoomRevenue] = null then 0 else [RoomRevenue],
                occ = if [RoomsOccupied] = null or [RoomsOccupied] = 0
                      then null else [RoomsOccupied]
            in
                if occ = null then null else rev / occ,
        type number
    ),

    AddRevPAR = Table.AddColumn(
        AddADR,
        "RevPAR",
        each
            let
                rev   = if [RoomRevenue] = null then 0 else [RoomRevenue],
                avail = if [RoomsAvailable] = null or [RoomsAvailable] = 0
                        then TotalRooms else [RoomsAvailable]
            in
                if avail = 0 then null else rev / avail,
        type number
    ),

    // ── Step 8: Select final columns ──
    SelectFinal = Table.SelectColumns(
        AddRevPAR,
        {
            "Date", "FileDate",
            "RoomsAvailable", "RoomsOccupied", "RoomsOOO",
            "RoomsComp", "RoomsHouseUse",
            "RoomRevenue", "FBRevenue", "OtherRevenue", "TotalRevenue",
            "OccupancyPct", "ADR", "RevPAR"
        },
        MissingField.Ignore
    ),

    // ── Step 9: Remove duplicates and sort ──
    RemoveDuplicates = Table.Distinct(SelectFinal, {"Date"}),
    SortByDate = Table.Sort(RemoveDuplicates, {{"Date", Order.Ascending}})

in
    SortByDate


──────────────────────────────────────────
QUERY: Fact_Reservations
Purpose: Loads and cleans Opera Reservation Activity exports.
         These have one row per reservation (or per reservation per night
         depending on how Opera is configured to export).

NOTE: Opera reservation exports vary significantly by property configuration.
      The column names below are commonly seen — adjust to match your client's actual export.
──────────────────────────────────────────

let
    // ── Step 1: Load files from reservations folder ──
    Source = Folder.Files(FolderPath_Reservations),
    
    ExcelFilesOnly = Table.SelectRows(
        Source,
        each Text.Lower(Text.End([Name], 5)) = ".xlsx"
           or Text.Lower(Text.End([Name], 4)) = ".xls"
    ),
    
    AddContent = Table.AddColumn(
        ExcelFilesOnly, "Data",
        each Excel.Workbook([Content], true, true)
    ),
    
    ExpandSheets = Table.ExpandTableColumn(
        AddContent, "Data",
        {"Name", "Data"}, {"SheetName", "SheetData"}
    ),
    
    FilterFirstSheet = Table.SelectRows(
        ExpandSheets,
        each [SheetName] = "Sheet1" or [SheetName] = "Report"
    ),
    
    ExpandData = Table.ExpandTableColumn(
        FilterFirstSheet, "SheetData",
        Table.ColumnNames(FilterFirstSheet{0}[SheetData])
    ),
    
    // ── Step 2: Skip Opera header rows and promote real headers ──
    SkipRows = 4,  // ← ADJUST per client
    RemoveTopRows = Table.Skip(ExpandData, SkipRows),
    PromoteHeaders = Table.PromoteHeaders(RemoveTopRows, [PromoteAllScalars = true]),
    
    // ── Step 3: Clean junk rows (use Confirmation Number as key — always populated for real reservations) ──
    CleanedRows = fnRemoveOperaJunk(PromoteHeaders, "Conf#"),  // ← Adjust column name if needed
    
    // ── Step 4: Rename columns to standard names ──
    // Adjust the left side values to match your client's Opera export exactly.
    RenameColumns = Table.RenameColumns(
        CleanedRows,
        {
            {"Conf#",                   "ConfirmationNumber"},
            {"Status",                  "ReservationStatus"},
            {"Arrival",                 "ArrivalDate_Raw"},
            {"Departure",               "DepartureDate_Raw"},
            {"Nights",                  "Nights_Raw"},
            {"Adults",                  "Adults_Raw"},
            {"Children",                "Children_Raw"},
            {"Room Type",               "RoomType"},
            {"Room Number",             "RoomNumber"},
            {"Rate Code",               "RateCode"},
            {"Rate Amount",             "RateAmount_Raw"},
            {"Market",                  "MarketSegment"},
            {"Source",                  "SourceOfBusiness"},
            {"Nationality",             "GuestNationality"},
            {"Booking Date",            "BookingDate_Raw"},
            {"Channel",                 "BookingChannel"},
            {"Company",                 "CompanyName"},
            {"Travel Agent",            "TravelAgent"},
            {"VIP",                     "VIPCode"}
        },
        MissingField.Ignore
    ),
    
    // ── Step 5: Parse dates ──
    ParseArrival = Table.AddColumn(
        RenameColumns, "ArrivalDate",
        each try fnParseOperaDate(Text.From([ArrivalDate_Raw]), DateFormat) otherwise null,
        type date
    ),
    
    ParseDeparture = Table.AddColumn(
        ParseArrival, "DepartureDate",
        each try fnParseOperaDate(Text.From([DepartureDate_Raw]), DateFormat) otherwise null,
        type date
    ),
    
    ParseBookingDate = Table.AddColumn(
        ParseDeparture, "BookingDate",
        each try fnParseOperaDate(Text.From([BookingDate_Raw]), DateFormat) otherwise null,
        type date
    ),
    
    // ── Step 6: Parse numerics ──
    ParseNumerics = Table.TransformColumns(
        ParseBookingDate,
        {
            {"Nights_Raw",      each fnCleanNumeric(_), type number},
            {"Adults_Raw",      each fnCleanNumeric(_), type number},
            {"Children_Raw",    each fnCleanNumeric(_), type number},
            {"RateAmount_Raw",  each fnCleanNumeric(_), type number}
        }
    ),
    
    // ── Step 7: Add calculated columns ──
    
    // Lead time in days (how far in advance was the booking made)
    AddLeadTime = Table.AddColumn(
        ParseNumerics, "BookingLeadDays",
        each if [BookingDate] = null or [ArrivalDate] = null 
             then null
             else Duration.Days([ArrivalDate] - [BookingDate]),
        type number
    ),
    
    // Total revenue per reservation
    AddTotalRate = Table.AddColumn(
        AddLeadTime, "TotalRateRevenue",
        each 
            let
                nights = if [Nights_Raw] = null then 0 else [Nights_Raw],
                rate   = if [RateAmount_Raw] = null then 0 else [RateAmount_Raw]
            in
                nights * rate,
        type number
    ),
    
    // Standardise reservation status to known values
    StandardiseStatus = Table.TransformColumns(
        AddTotalRate,
        {
            {"ReservationStatus", 
             each 
                if Text.Contains(Text.Upper(Text.From(_)), "CANCEL") then "Cancelled"
                else if Text.Contains(Text.Upper(Text.From(_)), "NO SHOW") then "No Show"
                else if Text.Contains(Text.Upper(Text.From(_)), "IN HOUSE") then "In House"
                else if Text.Contains(Text.Upper(Text.From(_)), "CHECKED OUT") then "Checked Out"
                else if Text.Contains(Text.Upper(Text.From(_)), "DUE IN") then "Due In"
                else if Text.Contains(Text.Upper(Text.From(_)), "RESERVED") then "Reserved"
                else Text.From(_),
             type text}
        }
    ),
    
    // ── Step 8: Select final columns ──
    SelectFinal = Table.SelectColumns(
        StandardiseStatus,
        {
            "ConfirmationNumber",
            "ReservationStatus",
            "ArrivalDate",
            "DepartureDate",
            "BookingDate",
            "Nights_Raw",
            "Adults_Raw",
            "Children_Raw",
            "RoomType",
            "RoomNumber",
            "RateCode",
            "RateAmount_Raw",
            "TotalRateRevenue",
            "BookingLeadDays",
            "MarketSegment",
            "SourceOfBusiness",
            "GuestNationality",
            "BookingChannel",
            "CompanyName",
            "TravelAgent",
            "VIPCode"
        },
        MissingField.Ignore
    ),
    
    FinalRename = Table.RenameColumns(
        SelectFinal,
        {
            {"Nights_Raw",      "Nights"},
            {"Adults_Raw",      "Adults"},
            {"Children_Raw",    "Children"},
            {"RateAmount_Raw",  "RateAmount"}
        },
        MissingField.Ignore
    ),
    
    RemoveDuplicates = Table.Distinct(FinalRename, {"ConfirmationNumber"}),
    SortByArrival = Table.Sort(RemoveDuplicates, {{"ArrivalDate", Order.Ascending}})

in
    SortByArrival


──────────────────────────────────────────
QUERY: Fact_SegmentStats
Purpose: Loads and cleans Opera Market Segment / Source of Business reports.
         Reads annual summary exports where each row is a segment and columns
         are monthly room nights (Jan–Dec) plus totals.
         Output is unpivoted to long format: one row per segment per month.

CHANGED in v1.1: Completely rewritten. Now uses an inline ProcessFile function,
         renames columns by position (more robust than by name), unpivots monthly
         RN columns to long format, and adds MonthNum for date relationship joins.
         No longer uses fnRemoveOperaJunk or fnCleanNumeric — filtering is inline.
         SkipRows is now hardcoded to 5 — adjust if your export has more/fewer header rows.
──────────────────────────────────────────

let
    Source = Folder.Files(FolderPath_Segments),

    ExcelFilesOnly = Table.SelectRows(
        Source,
        each Text.Lower(Text.End([Name], 5)) = ".xlsx"
    ),

    // ── Per-file processing function ──
    ProcessFile = (fileContent as binary) as table =>
        let
            Workbook = Excel.Workbook(fileContent, true, true),

            // Get the Report sheet, fall back to first sheet
            Sheet =
                let
                    matching = Table.SelectRows(Workbook, each [Name] = "Report")
                in
                    if Table.RowCount(matching) > 0
                    then matching{0}[Data]
                    else Workbook{0}[Data],

            // Skip Opera header rows (adjust 5 if your export differs)
            Skipped = Table.Skip(Sheet, 5),
            Promoted = Table.PromoteHeaders(Skipped, [PromoteAllScalars = true]),

            // Rename columns by position — more robust than by name when Opera
            // column headers vary by property config.
            // Expected column order: Segment, Group, Jan..Dec, Total RN, Total Rev, ADR, Mix%
            RenamedByPosition = Table.RenameColumns(
                Promoted,
                List.Zip({
                    Table.ColumnNames(Promoted),
                    {"MarketSegment", "SegmentGroup",
                     "RN_Jan", "RN_Feb", "RN_Mar", "RN_Apr",
                     "RN_May", "RN_Jun", "RN_Jul", "RN_Aug",
                     "RN_Sep", "RN_Oct", "RN_Nov", "RN_Dec",
                     "TotalRoomNights", "TotalRevenue", "ADR", "RevenueMixPct"}
                })
            ),

            // Remove junk rows inline (no helper function needed)
            CleanedRows = Table.SelectRows(
                RenamedByPosition,
                each [MarketSegment] <> null
                     and [MarketSegment] <> ""
                     and [MarketSegment] <> "GRAND TOTAL"
                     and [MarketSegment] <> "*** End of Report ***"
            ),

            // Set types on summary and monthly columns
            WithTypes = Table.TransformColumnTypes(
                CleanedRows,
                {
                    {"TotalRoomNights", type number},
                    {"TotalRevenue",    type number},
                    {"ADR",             type number},
                    {"RevenueMixPct",   type number},
                    {"RN_Jan",  type number}, {"RN_Feb",  type number},
                    {"RN_Mar",  type number}, {"RN_Apr",  type number},
                    {"RN_May",  type number}, {"RN_Jun",  type number},
                    {"RN_Jul",  type number}, {"RN_Aug",  type number},
                    {"RN_Sep",  type number}, {"RN_Oct",  type number},
                    {"RN_Nov",  type number}, {"RN_Dec",  type number}
                }
            )
        in
            WithTypes,

    // Apply function to every file
    AddProcessed = Table.AddColumn(
        ExcelFilesOnly,
        "Processed",
        each ProcessFile([Content])
    ),

    // Expand all columns from the processed tables
    Expanded = Table.ExpandTableColumn(
        AddProcessed,
        "Processed",
        {"MarketSegment", "SegmentGroup",
         "RN_Jan", "RN_Feb", "RN_Mar", "RN_Apr",
         "RN_May", "RN_Jun", "RN_Jul", "RN_Aug",
         "RN_Sep", "RN_Oct", "RN_Nov", "RN_Dec",
         "TotalRoomNights", "TotalRevenue", "ADR", "RevenueMixPct"},
        {"MarketSegment", "SegmentGroup",
         "RN_Jan", "RN_Feb", "RN_Mar", "RN_Apr",
         "RN_May", "RN_Jun", "RN_Jul", "RN_Aug",
         "RN_Sep", "RN_Oct", "RN_Nov", "RN_Dec",
         "TotalRoomNights", "TotalRevenue", "ADR", "RevenueMixPct"}
    ),

    // Unpivot monthly RN columns → long format (one row per segment per month)
    Unpivoted = Table.UnpivotOtherColumns(
        Expanded,
        {"MarketSegment", "SegmentGroup",
         "TotalRoomNights", "TotalRevenue", "ADR", "RevenueMixPct"},
        "MonthCol", "RoomNights"
    ),

    // Convert MonthCol (e.g. "RN_Jan") to a numeric month number for date joins
    AddMonthNum = Table.AddColumn(
        Unpivoted,
        "MonthNum",
        each List.PositionOf(
            {"RN_Jan","RN_Feb","RN_Mar","RN_Apr",
             "RN_May","RN_Jun","RN_Jul","RN_Aug",
             "RN_Sep","RN_Oct","RN_Nov","RN_Dec"},
            [MonthCol]
        ) + 1,
        Int64.Type
    ),

    // Remove any rows where the month lookup didn't match (shouldn't happen)
    FilterValidMonths = Table.SelectRows(
        AddMonthNum,
        each [MonthNum] > 0
    ),

    DropMonthCol = Table.RemoveColumns(FilterValidMonths, {"MonthCol"}),

    // Recalculate ADR from totals (more reliable than Opera's exported ADR)
    AddCalcADR = Table.AddColumn(
        DropMonthCol,
        "ADR_Calculated",
        each if [RoomNights] = null or [RoomNights] = 0
             then null
             else [TotalRevenue] / [TotalRoomNights],
        type number
    ),

    // Select and order final columns
    SelectFinal = Table.SelectColumns(
        AddCalcADR,
        {"MarketSegment", "SegmentGroup",
         "MonthNum", "RoomNights",
         "TotalRoomNights", "TotalRevenue",
         "ADR", "ADR_Calculated", "RevenueMixPct"},
        MissingField.Ignore
    ),

    SortFinal = Table.Sort(
        SelectFinal,
        {{"MarketSegment", Order.Ascending}, {"MonthNum", Order.Ascending}}
    )

in
    SortFinal


──────────────────────────────────────────
QUERY: Dim_MarketSegmentMapping
Purpose: Maps Opera's raw segment codes (which can be 20-30 granular codes)
         to your 6 standard reporting groups.
         This is a manually maintained table — you set it up once per client.

HOW TO USE:
  - Create this as a Blank Query
  - Or create it as an Excel file the client can edit (recommended — lets them
    reclassify segments without calling you)
──────────────────────────────────────────

let
    // Define your segment mapping as an inline table
    // Left column = exactly as it appears in Opera export
    // Right column = your standard reporting group
    
    SegmentMapping = #table(
        {"OperaSegmentCode", "SegmentGroup", "SegmentDescription", "SortOrder"},
        {
            // OTA
            {"BOOKINGCOM",      "OTA",          "Booking.com",              1},
            {"EXPEDIA",         "OTA",          "Expedia",                  1},
            {"AIRBNB",          "OTA",          "Airbnb",                   1},
            {"OTA",             "OTA",          "Online Travel Agent",       1},
            {"HOTELSCOM",       "OTA",          "Hotels.com",               1},
            
            // Direct
            {"DIRECT",          "Direct",       "Direct Booking",           2},
            {"WEBSITE",         "Direct",       "Hotel Website",            2},
            {"WALKIN",          "Direct",       "Walk In",                  2},
            {"PHONE",           "Direct",       "Phone Reservation",        2},
            {"EMAIL",           "Direct",       "Email Reservation",        2},
            
            // Corporate
            {"CORPORATE",       "Corporate",    "Corporate Rate",           3},
            {"CORP",            "Corporate",    "Corporate",                3},
            {"NEGOTIATED",      "Corporate",    "Negotiated Rate",          3},
            {"BUSINESS",        "Corporate",    "Business Traveller",       3},
            
            // Groups & Events
            {"GROUP",           "Groups",       "Group Booking",            4},
            {"MICE",            "Groups",       "MICE/Events",              4},
            {"WEDDING",         "Groups",       "Wedding Group",            4},
            {"EVENT",           "Groups",       "Event Group",              4},
            
            // Wholesale & Tour Operators
            {"WHOLESALE",       "Wholesale",    "Wholesale",                5},
            {"TOUROPERATOR",    "Wholesale",    "Tour Operator",            5},
            {"TO",              "Wholesale",    "Tour Operator",            5},
            {"PACKAGE",         "Wholesale",    "Package",                  5},
            
            // Other / Complimentary
            {"COMP",            "Other",        "Complimentary",            6},
            {"STAFF",           "Other",        "Staff Rate",               6},
            {"HOUSEUSE",        "Other",        "House Use",                6},
            {"OTHER",           "Other",        "Other",                    6}
        }
    )
in
    SegmentMapping


──────────────────────────────────────────
QUERY: DimDate  (named "Dim_Date" in template v1.0 — renamed to DimDate in the model)
Purpose: Standard date dimension table.
         Built entirely in Power Query — no dependency on Opera exports.
         Respects the FiscalYearStartMonth parameter.

NOTE: For a hotel, the most important date flags are:
  - Day of week (weekday vs weekend pricing)
  - Season (peak / shoulder / low — configure in Dim_Season separately)
  - School holidays (if your hotel targets families)
──────────────────────────────────────────

let
    // ── Configure date range ──
    StartDate = #date(2022, 1, 1),   // Set to 2 years before your data starts
    EndDate   = #date(2027, 12, 31), // Set to 2 years beyond current year
    
    // ── Generate date list ──
    DateList = List.Dates(
        StartDate,
        Duration.Days(EndDate - StartDate) + 1,
        #duration(1, 0, 0, 0)
    ),
    
    DateTable = Table.FromList(DateList, Splitter.SplitByNothing(), {"Date"}),
    TypeDate = Table.TransformColumnTypes(DateTable, {{"Date", type date}}),
    
    // ── Add date attributes ──
    AddYear = Table.AddColumn(TypeDate, "Year", each Date.Year([Date]), Int64.Type),
    AddMonth = Table.AddColumn(AddYear, "MonthNumber", each Date.Month([Date]), Int64.Type),
    AddMonthName = Table.AddColumn(AddMonth, "MonthName", each Date.ToText([Date], "MMMM"), type text),
    AddMonthShort = Table.AddColumn(AddMonthName, "MonthShort", each Date.ToText([Date], "MMM"), type text),
    AddQuarter = Table.AddColumn(AddMonthShort, "Quarter", each "Q" & Text.From(Date.QuarterOfYear([Date])), type text),
    AddQuarterNum = Table.AddColumn(AddQuarter, "QuarterNumber", each Date.QuarterOfYear([Date]), Int64.Type),
    AddDayOfWeek = Table.AddColumn(AddQuarterNum, "DayOfWeek", each Date.DayOfWeek([Date], Day.Monday) + 1, Int64.Type),
    AddDayName = Table.AddColumn(AddDayOfWeek, "DayName", each Date.ToText([Date], "dddd"), type text),
    AddDayShort = Table.AddColumn(AddDayName, "DayShort", each Date.ToText([Date], "ddd"), type text),
    AddWeekNumber = Table.AddColumn(AddDayShort, "WeekNumber", each Date.WeekOfYear([Date], Day.Monday), Int64.Type),
    AddWeekYear = Table.AddColumn(AddWeekNumber, "WeekYear", each Text.From([Year]) & "-W" & Text.PadStart(Text.From([WeekNumber]), 2, "0"), type text),
    
    // Is weekend flag (Friday + Saturday nights are the key nights for hotels)
    // Adjust if your market has different weekend patterns
    AddIsWeekend = Table.AddColumn(
        AddWeekYear, "IsWeekend",
        each if [DayOfWeek] >= 5 then true else false,  // 5=Fri, 6=Sat, 7=Sun in Mon-start week
        type logical
    ),
    
    // Month-Year for grouping in charts
    AddMonthYear = Table.AddColumn(
        AddIsWeekend, "MonthYear",
        each Date.ToText([Date], "MMM yyyy"),
        type text
    ),
    
    // Month-Year sort key (numeric, so charts sort correctly)
    AddMonthYearSort = Table.AddColumn(
        AddMonthYear, "MonthYearSort",
        each [Year] * 100 + [MonthNumber],
        Int64.Type
    ),
    
    // Fiscal year (respects FiscalYearStartMonth parameter)
    AddFiscalYear = Table.AddColumn(
        AddMonthYearSort, "FiscalYear",
        each 
            if [MonthNumber] >= FiscalYearStartMonth 
            then [Year]
            else [Year] - 1,
        Int64.Type
    ),
    
    AddFiscalYearLabel = Table.AddColumn(
        AddFiscalYear, "FiscalYearLabel",
        each "FY" & Text.From([FiscalYear]),
        type text
    ),
    
    // Fiscal month (month 1 = first month of fiscal year)
    AddFiscalMonth = Table.AddColumn(
        AddFiscalYearLabel, "FiscalMonth",
        each 
            let m = [MonthNumber] - FiscalYearStartMonth + 1
            in if m <= 0 then m + 12 else m,
        Int64.Type
    ),
    
    // Is past or future (useful for conditional formatting in reports)
    AddIsPast = Table.AddColumn(
        AddFiscalMonth, "IsPast",
        each if [Date] < Date.From(DateTime.LocalNow()) then true else false,
        type logical
    ),
    
    AddIsToday = Table.AddColumn(
        AddIsPast, "IsToday",
        each [Date] = Date.From(DateTime.LocalNow()),
        type logical
    ),
    
    // Date key for joining (YYYYMMDD integer — useful for relationships)
    AddDateKey = Table.AddColumn(
        AddIsToday, "DateKey",
        each [Year] * 10000 + [MonthNumber] * 100 + Date.Day([Date]),
        Int64.Type
    ),
    
    SortByDate = Table.Sort(AddDateKey, {{"Date", Order.Ascending}})

in
    SortByDate


═══════════════════════════════════════════════════════════════
SECTION 4 — SETUP CHECKLIST
(Work through this for every new client onboarding)
═══════════════════════════════════════════════════════════════

□ 1. Collect 3–6 months of sample Opera exports from the client
     Ask for: Daily Flash, Reservation Activity, Market Segment report

□ 2. Open each export in Excel and note:
     - How many rows before the real header? (sets SkipRows value)
     - Exact column names for dates, rooms, revenue
     - Date format (DD/MM/YYYY or MM/DD/YYYY)
     - Are numbers formatted as text with currency/% symbols?
     - Are there sub-total rows within the data?

□ 3. Update Parameters:
     - HotelName, TotalRooms, FiscalYearStartMonth, CurrencySymbol, DateFormat
     - FolderPath values

□ 4. Update Fact_DailyFlash query:
     - SkipRows value
     - KeyColumnName (the date column's exact name)
     - RenameColumns pairs (left side = Opera names, right side = standard names)

□ 5. Update Fact_Reservations query:
     - Same SkipRows and column rename adjustments
     - Adjust the key column in fnRemoveOperaJunk if "Conf#" differs

□ 6. Update Dim_MarketSegmentMapping:
     - Add/remove rows to match client's actual Opera segment codes
     - Run a quick DISTINCT on the MarketSegment column from raw data to get the full list

□ 7. Verify Dim_Date:
     - Confirm FiscalYearStartMonth is set correctly in Parameters
     - Confirm date range covers all historical data plus 2 future years

□ 8. Test refresh end-to-end:
     - Drop a new Flash export into the folder
     - Hit Refresh in Power BI
     - Verify the new date appears in Fact_DailyFlash
     - Verify no duplicates on the Date column

□ 9. Set up scheduled refresh (if client is on Power BI Pro + gateway):
     - Install On-Premises Data Gateway on the hotel's server
     - Configure folder data source in Power BI Service
     - Set refresh schedule (daily, after their morning Opera export)

□ 10. Document the SOP for the hotel:
      - Which Opera reports to run and when
      - File naming convention
      - Which folder to drop files into
      - What to do if refresh fails (email you)


═══════════════════════════════════════════════════════════════
SECTION 5 — COMMON TROUBLESHOOTING
═══════════════════════════════════════════════════════════════

PROBLEM: "Expression.Error: The column X was not found"
FIX: Opera column name in the export doesn't match your RenameColumns pair.
     Open the raw file, check exact column name including spaces and capitalisation.
     Or set MissingField.Ignore on the rename step and check what comes through.

PROBLEM: Dates showing as null after parsing
FIX: The date format doesn't match the DateFormat parameter.
     Open a raw file, look at an actual date cell value.
     Switch DateFormat parameter between "DMY" and "MDY".

PROBLEM: Revenue columns showing null
FIX: Opera has added a character to the number that fnCleanNumeric doesn't strip.
     Check the raw text value. Common culprits: non-breaking space (char 160),
     special dash characters, locale-specific decimal separators (comma vs period).
     Add the offending character to the Text.Replace chain in fnCleanNumeric.

PROBLEM: Totals rows appearing in the data
FIX: The totals row uses a different string than fnRemoveOperaJunk expects.
     Add that string to the junkStrings list in fnRemoveOperaJunk.

PROBLEM: Duplicate dates in Fact_DailyFlash
FIX: Same file dropped twice, or Opera exported two files for the same date.
     The Table.Distinct step at the end handles this. If you need both versions
     (to spot discrepancies), remove the Distinct step and add a FileDate column.

PROBLEM: "DataFormat.Error: We couldn't convert to Number"
FIX: A row that should have been removed by fnRemoveOperaJunk is still in the data.
     Find the offending row by removing the final type conversion step temporarily.
     Add its distinguishing text to the junkStrings list.

PROBLEM: Report refresh works locally but fails in Power BI Service
FIX: The gateway can't access the folder path. Confirm:
     - The gateway machine can see the folder (it's on a shared drive or local to the gateway machine)
     - The folder path in the Parameter matches exactly what the gateway machine sees
     - The gateway service account has read permissions on the folder


═══════════════════════════════════════════════════════════════
CHANGELOG
═══════════════════════════════════════════════════════════════

v1.1 — Updated from hotelbimodel.bim
  • Fact_SegmentStats — fully rewritten:
      - Now uses an inline ProcessFile function (consistent pattern with Fact_DailyFlash)
      - Columns renamed by position instead of by name (more robust across Opera configs)
      - Output unpivoted to long format: one row per segment per month (MonthNum 1–12)
      - Monthly RN columns: RN_Jan through RN_Dec, then unpivoted to RoomNights + MonthNum
      - Inline junk-row filtering replaces fnRemoveOperaJunk dependency
      - SkipRows changed from 4 to 5 — verify against your client's actual export
      - No longer outputs Year/Month/Period columns — use MonthNum + join to DimDate instead
  • DimDate — renamed from Dim_Date. M code unchanged.
  • All other queries, functions and parameters — unchanged from v1.0.

v1.0 — Initial template
═══════════════════════════════════════════════════════════════
END OF TEMPLATE
Opera Insights Power Query Template v1.1
© [Your Name] Consulting
═══════════════════════════════════════════════════════════════
