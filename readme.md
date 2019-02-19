# Tools to pull data out of SpiceWorks database.

## "spiceworks_weekly_report.ps1"

Pulls all open tickets and all closed within the last week and sends report to a physical printer, incluidng ASCII-Art goodness

## "ticket_display_to_wwwserver.ps1"

Pulls open tickets, does some replacey-magic on "Tech ID" to match it to actual technician names, and then creates a webpage of the results.  Includes a generic Spiceworks-like css formatting.

## "ticket_display_to_wwwserver-terminal-style.ps1"

Same as above, but uses some Archer-inspired terminal design.