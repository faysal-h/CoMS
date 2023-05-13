
# CCMS
A project to remove removing redundant data entry and automate comparison case worksheets and report generator.

## Dependencies
Following list of modules are necessary to build the executable
1. python-dateutil
2. inflect
3. sqlalchemy_access
4. pandas
5. docxtpl
6. pymsgbox

## Documentation
User must enter details of a case in ACCESS database file CCMSDatabas. After entering the details run CCMS from the main folder. following options are available
1. Generate sheets
This option generates sheets and reports of a single case.
2. Generate Batch sheets
This option generates all sheets and reports of all cases in a database
3. Generate Identifiers
This option generates Case Identifiers, CPR list, and Envelops Identifiers
