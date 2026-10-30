# Groups the database tables into the tabs people see in the snapshot.
#
# - Each category is a top tab. With Placement = 'Sidebar', its tabs show as
#   buttons in the left sidebar instead.
# - Each tab shows the table whose NAME matches Match. * means "anything", so
#   '*Application*' matches tblApplications, Applications_List, and so on.
#   Match can be a list: Match = '*Token*', '*OTP*'
# - A table goes to the FIRST tab that matches it, so put specific patterns
#   (like '*Non*GIAM*') before general ones (like '*GIAM*').
# - If a pattern matches several tables, each gets its own tab with its
#   table name. Categories that match nothing are hidden.
# - Tables that match nothing show under "Other Tables"; saved queries show
#   under "Queries". Nothing is ever left out.
#
# The export prints which table went where, so adjust the patterns after the
# first run if something lands in the wrong place.

@{
    Categories = @(
        @{
            Name      = 'Contacts'
            Placement = 'Sidebar'
            Tabs      = @(
                @{ Label = 'Non-GIAM Contacts'; Match = '*Non*GIAM*' }
                @{ Label = 'GIAM Contacts'; Match = '*GIAM*' }
                @{ Label = 'Executive Support'; Match = '*Exec*' }
                @{ Label = 'SC Phone Menu'; Match = '*Phone*' }
                @{ Label = 'SecOps Teams'; Match = '*SecOps*Team*' }
            )
        }
        @{
            Name = 'Search Tools'
            Tabs = @(
                @{ Label = 'Applications'; Match = '*Application*' }
                @{ Label = 'SecOPS'; Match = '*SecOps*' }
                @{ Label = 'ISA Contacts'; Match = '*ISA*Contact*' }
                @{ Label = 'RIT Groups'; Match = '*RIT*Group*' }
            )
        }
        @{
            Name = 'Token Tools'
            Tabs = @(
                @{ Label = 'Tokens'; Match = '*Token*' }
            )
        }
        @{
            Name = 'Mainframe Tools'
            Tabs = @(
                @{ Label = 'Mainframe'; Match = '*Mainframe*' }
            )
        }
        # The two groups below sit in the sidebar under the contacts, as in HELP+
        @{
            Name      = 'Tracking tools'
            Placement = 'Sidebar'
            Tabs      = @(
                @{ Label = 'Severity'; Match = '*Severity*' }
                @{ Label = 'MIM Tracker'; Match = '*MIM*' }
                @{ Label = 'INC Tracker'; Match = '*INC*Track*', '*Incident*' }
            )
        }
        @{
            Name      = 'Copy tools'
            Placement = 'Sidebar'
            Tabs      = @(
                @{ Label = 'INC groups'; Match = '*INC*Group*' }
                @{ Label = 'Canned responses'; Match = '*Canned*', '*Response*' }
            )
        }
    )
}
