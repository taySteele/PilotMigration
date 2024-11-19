

# Pilot Provisioning

To pilot the provisioning fo ~100 sites of different types, and business units.

## Pre-requisites

* Foundation buid completed.
* Retension policies turned off in purview for all Pilot sites.

## Provisioning Development Tasks

Automated provisioning of ~100 sites from a csv file.

* For each source site, create new destination site
* Use site type e.g. Communications
* Apply site template (Pnp xml) e.g. Community Site Template
* Apply site settings
* Activate site\site collection features
* Copy lists and libraries. (Sharegate PnP)
- Apply custom columns. Hidden from users but added to all lists\libraries. Populated during migration. Do we need to add this to all content types on the list?
    - COHESION-LEGACY-ID - Single Line Text
    - COHESION-LEGACY-URI - Multi Line 
    - What other custom fields do we need? 
* Ensure retension polcies are turned off within all libraries. (Purview setting managed separately.)
* Teamify the site.

## Non functional requirements

* Log output of actions performed.
* Log output from ShareGate. (Info, Warning, Errors) - log out to csv for all commands used.
* Record times for performance indicators.
* Log time to provision a new site.
* Log time to create a document library.

e..g 

## Define log.csv columns

Destination site name
Destination site address
Log Level (Info, Warning, Error)
Log Message
Log Time






