"""
Procurement Categorization Engine — 5-Pass Rule Stack
Built from VCU's actual commodity_codes.xlsx and account_codes.csv
"""
import re
import pandas as pd

BUCKET_COLORS = {
    "Research & Laboratory":                "#3B82F6",
    "Clinical / Healthcare":                "#EC4899",
    "IT":                                   "#F59E0B",
    "Facilities / MRO":                     "#84CC16",
    "Utilities & Occupancy":                "#10B981",
    "Admin & Office":                       "#6B7280",
    "Travel, Events & Hospitality":         "#14B8A6",
    "Services":                             "#8B5CF6",
    "Capital Projects & Construction":      "#A855F7",
    "Printing, Marketing & Communications": "#F97316",
    "Food & Catering":                      "#EF4444",
    "Inter-Entity / Transfers":             "#64748B",
    "Uncategorized":                        "#374151",
}

# ══════════════════════════════════════════════════════════════
# PASS 1 — COMMODITY CODE CROSSWALK
# Source: VCU commodity_codes.xlsx (316 codes)
# ══════════════════════════════════════════════════════════════
COMMODITY_CROSSWALK = {
    # ── CONSTRUCTION (CON) ────────────────────────────────────
    "721030": ("Facilities / MRO","Trades Services","Site Preparation"),
    "721317": ("Facilities / MRO","Trades Services","Infrastructure"),
    "721400": ("Capital Projects & Construction","Architecture & Engineering","A&E Design Services"),
    "721500": ("Facilities / MRO","Trades Services","Specialized Trade Construction"),
    "721527": ("Facilities / MRO","Trades Services","Concrete Work"),
    "950000": ("Facilities / MRO","Trades Services","Land & Structures"),

    # ── EQUIPMENT (EQUIP) ─────────────────────────────────────
    "28700":  ("IT","IT Hardware & Peripherals","Electronic Equipment"),
    "101300": ("Research & Laboratory","Lab Equipment & Instruments","Animal Containment"),
    "221000": ("Facilities / MRO","MRO Supplies","Heavy Construction Machinery"),
    "231800": ("Food & Catering","Food & Beverage","Industrial Food Equipment"),
    "241000": ("Facilities / MRO","MRO Supplies","Material Handling Equipment"),
    "241300": ("Facilities / MRO","MRO Supplies","Industrial Refrigeration"),
    "251000": ("Facilities / MRO","MRO Supplies","Motor Vehicles"),
    "251700": ("Facilities / MRO","MRO Supplies","Vehicle Components"),
    "270000": ("Facilities / MRO","MRO Supplies","Tools & General Machinery"),
    "271100": ("Facilities / MRO","MRO Supplies","Hand Tools"),
    "302300": ("Facilities / MRO","Trades Services","Portable Structures"),
    "302315": ("Travel, Events & Hospitality","Event / Venue","Grandstands & Bleachers"),
    "302317": ("Travel, Events & Hospitality","Event / Venue","Tents & Temporary Structures"),
    "391200": ("Facilities / MRO","MRO Supplies","Electrical Equipment"),
    "410000": ("Research & Laboratory","Lab Equipment & Instruments","Lab Measuring & Testing Equipment"),
    "411000": ("Research & Laboratory","Lab Equipment & Instruments","Lab & Scientific Equipment"),
    "411100": ("Research & Laboratory","Lab Equipment & Instruments","Measuring & Testing Instruments"),
    "420000": ("Clinical / Healthcare","Medical Equipment","Medical Equipment & Supplies"),
    "421500": ("Clinical / Healthcare","Medical Equipment","Dental Equipment & Supplies"),
    "431100": ("IT","IT Hardware & Peripherals","Computer Hardware Drives"),
    "431310": ("IT","IT Hardware & Peripherals","Laptops & Tablets"),
    "431320": ("IT","IT Hardware & Peripherals","Desktops & Monitors"),
    "431332": ("IT","IT Hardware & Peripherals","Scanners"),
    "432100": ("IT","IT Hardware & Peripherals","Computers"),
    "432119": ("IT","IT Hardware & Peripherals","Computer Displays"),
    "432121": ("IT","IT Hardware & Peripherals","Computer Printers"),
    "432200": ("IT","Network / Telecom","Network Equipment & Platforms"),
    "432217": ("IT","Network / Telecom","Fixed Network Equipment"),
    "432218": ("IT","Network / Telecom","Optical Network Devices"),
    "432225": ("IT","IT Software / SaaS","Network Security Equipment"),
    "432226": ("IT","IT Services","Network Service"),
    "432229": ("IT","Network / Telecom","Telephony Equipment"),
    "432231": ("IT","Network / Telecom","Mobile Network Infrastructure"),
    "432233": ("IT","Network / Telecom","Datacom & Connectivity Devices"),
    "434100": ("IT","IT Hardware & Peripherals","Servers"),
    "440000": ("Admin & Office","Furniture & Fixtures","Office Equipment & Desks"),
    "441000": ("Admin & Office","Office Machines","Copiers & Office Machines"),
    "441015": ("Admin & Office","Office Machines","Photocopiers"),
    "451000": ("Printing, Marketing & Communications","Printing","Printing & Publishing Equipment"),
    "451015": ("Printing, Marketing & Communications","Printing","Printing Machinery"),
    "451100": ("IT","IT Hardware & Peripherals","AV Presentation & Conferencing Equipment"),
    "451115": ("IT","IT Hardware & Peripherals","Lecterns & Sound Systems"),
    "451116": ("IT","IT Hardware & Peripherals","Projectors & Supplies"),
    "451119": ("IT","IT Hardware & Peripherals","Video Conference Hardware"),
    "451120": ("Admin & Office","Published / Books","Microfilm Equipment"),
    "451200": ("Printing, Marketing & Communications","Marketing / Advertising","Photographic & Video Equipment"),
    "451215": ("Printing, Marketing & Communications","Marketing / Advertising","Cameras"),
    "451217": ("Printing, Marketing & Communications","Marketing / Advertising","Photographic Processing"),
    "451300": ("Printing, Marketing & Communications","Marketing / Advertising","Photographic & Recording Media"),
    "451317": ("IT","IT Hardware & Peripherals","Media Storage"),
    "461500": ("Admin & Office","General Admin","Law Enforcement Equipment"),
    "461700": ("Admin & Office","General Admin","Security Surveillance & Detection"),
    "461716": ("Admin & Office","General Admin","Surveillance & Detection"),
    "461825": ("Admin & Office","General Admin","Personal Safety Devices"),
    "461900": ("Facilities / MRO","MRO Supplies","Fire Protection"),
    "461915": ("Facilities / MRO","MRO Supplies","Fire Prevention & Suppression"),
    "471000": ("Utilities & Occupancy","Water / Sewer","Water Treatment Equipment"),
    "471200": ("Facilities / MRO","Janitorial","Janitorial Equipment"),
    "480000": ("Facilities / MRO","MRO Supplies","Service Industry Machinery"),
    "481000": ("Food & Catering","Food & Beverage","Food Service Equipment"),
    "481100": ("Food & Catering","Food & Beverage","Vending Machines"),
    "490000": ("Travel, Events & Hospitality","Event / Venue","Recreational Equipment"),
    "492000": ("Travel, Events & Hospitality","Event / Venue","Fitness Equipment"),
    "492200": ("Travel, Events & Hospitality","Event / Venue","Sports Equipment"),
    "551200": ("Admin & Office","General Admin","Signage Equipment"),
    "561220": ("Research & Laboratory","Lab Equipment & Instruments","Laboratory Furniture"),
    "601200": ("Admin & Office","General Admin","Instructional Arts & Crafts Equipment"),
    "601300": ("Admin & Office","General Admin","Musical Instruments"),

    # ── OTHER ─────────────────────────────────────────────────
    "101000": ("Research & Laboratory","Lab Equipment & Instruments","Animals"),
    "941000": ("Inter-Entity / Transfers","Internal Transfer","Work-Related Organizations"),
    "999094": ("Uncategorized","Unknown","Not Elsewhere Classified"),

    # ── SOFTWARE ──────────────────────────────────────────────
    "432232": ("IT","IT Software / SaaS","Mobile Messaging Platforms"),
    "432300": ("IT","IT Software / SaaS","Software - Other"),
    "432315": ("IT","IT Software / SaaS","Business Software"),
    "432316": ("IT","IT Software / SaaS","Finance / ERP Software"),
    "432320": ("IT","IT Software / SaaS","Computer Game / Entertainment Software"),
    "432322": ("IT","IT Software / SaaS","Content Management Software"),
    "432323": ("IT","IT Software / SaaS","Data Management & Query Software"),
    "432324": ("IT","IT Software / SaaS","Development Software"),
    "432325": ("IT","IT Software / SaaS","Educational / Reference Software"),
    "432327": ("IT","IT Software / SaaS","Network Applications & Management Software"),
    "432329": ("IT","IT Software / SaaS","Networking Software"),
    "432330": ("IT","IT Software / SaaS","Operating Environment Software"),
    "432332": ("IT","IT Software / SaaS","Security & Protection Software"),
    "432334": ("IT","IT Software / SaaS","Utility & Device Driver Software"),
    "432337": ("IT","IT Software / SaaS","System Management Software"),
    "433100": ("IT","IT Software / SaaS","Research/Academic Licensing"),
    "433200": ("IT","IT Software / SaaS","Enterprise Software (Banner etc)"),
    "441120": ("IT","IT Software / SaaS","Planning Software"),

    # ── SUPPLIES (SUP) ────────────────────────────────────────
    "101100": ("Research & Laboratory","Lab Consumables","Animal Products & Supplies"),
    "101600": ("Facilities / MRO","MRO Supplies","Horticulture & Landscape"),
    "101700": ("Facilities / MRO","MRO Supplies","Fertilizers & Herbicides"),
    "111600": ("Admin & Office","General Admin","Fabrics & Textiles"),
    "121400": ("Research & Laboratory","Chemicals & Reagents","Gases & Elements"),
    "141100": ("Admin & Office","Office Supplies","Paper Products"),
    "141114": ("Admin & Office","Office Supplies","Copy Paper"),
    "141115": ("Admin & Office","Office Supplies","Gift Cards & Cash Incentives"),
    "141116": ("Admin & Office","Office Supplies","Business Cards"),
    "141117": ("Admin & Office","Office Supplies","Paper Products - Personal"),
    "141118": ("Admin & Office","Office Supplies","Paper Products - Business"),
    "141200": ("Admin & Office","Office Supplies","Industrial Paper"),
    "151000": ("Facilities / MRO","MRO Supplies","Fuels"),
    "151200": ("Facilities / MRO","MRO Supplies","Lubricants & Oils"),
    "241400": ("Facilities / MRO","MRO Supplies","Packing Supplies"),
    "250000": ("Facilities / MRO","MRO Supplies","Vehicle Accessories & Components"),
    "261200": ("Facilities / MRO","MRO Supplies","Electrical Wire & Cable"),
    "301300": ("Facilities / MRO","MRO Supplies","Structural Building Materials"),
    "301500": ("Facilities / MRO","MRO Supplies","Exterior Finishing Materials"),
    "301600": ("Facilities / MRO","MRO Supplies","Interior Finishing Materials"),
    "311600": ("Facilities / MRO","MRO Supplies","Hardware"),
    "311700": ("Facilities / MRO","MRO Supplies","Bearings & Gears"),
    "312000": ("Facilities / MRO","MRO Supplies","Adhesives & Sealants"),
    "312100": ("Facilities / MRO","MRO Supplies","Paints & Finishes"),
    "391000": ("Facilities / MRO","MRO Supplies","Lamps & Lighting"),
    "391100": ("Facilities / MRO","MRO Supplies","Lighting Fixtures"),
    "401000": ("Facilities / MRO","MRO Supplies","HVAC Supplies"),
    "401615": ("Facilities / MRO","MRO Supplies","Filters"),
    "411200": ("Research & Laboratory","Lab Consumables","Laboratory Supplies"),
    "421400": ("Clinical / Healthcare","Medical Supplies","Patient Care & Treatment Supplies"),
    "421900": ("Clinical / Healthcare","Medical Supplies","Medical Facility Products"),
    "423000": ("Clinical / Healthcare","Medical Supplies","Medical Training Supplies"),
    "431340": ("IT","IT Hardware & Peripherals","PC/Laptop Accessories"),
    "432010": ("IT","IT Hardware & Peripherals","IT Storage - Memory/USB"),
    "432116": ("IT","IT Hardware & Peripherals","Computer Accessories"),
    "432117": ("IT","IT Hardware & Peripherals","Computer Input Devices"),
    "432118": ("IT","IT Hardware & Peripherals","Input Device Accessories"),
    "432120": ("IT","IT Hardware & Peripherals","Display Accessories"),
    "432215": ("IT","Network / Telecom","Call Management Systems"),
    "441100": ("Admin & Office","Office Supplies","Office & Desk Accessories"),
    "441200": ("Admin & Office","Office Supplies","Office Supplies - General"),
    "441210": ("Admin & Office","Office Supplies","Printer/Copier Ink & Toner"),
    "441214": ("Admin & Office","Office Supplies","Mailing Supplies"),
    "441215": ("Admin & Office","Office Supplies","Postage & Stamps"),
    "441216": ("Admin & Office","Office Supplies","Desk Supplies"),
    "441217": ("Admin & Office","Office Supplies","Writing Instruments"),
    "441220": ("Admin & Office","Office Supplies","Folders & Binders"),
    "451016": ("Printing, Marketing & Communications","Printing","Printing Machinery Accessories"),
    "451017": ("Admin & Office","Promotional / Branded Items","Printing Accessories"),
    "461600": ("Admin & Office","General Admin","Public Safety Supplies"),
    "461715": ("Admin & Office","General Admin","Locks & Security Hardware"),
    "461800": ("Admin & Office","General Admin","Personal Safety & Protection"),
    "461815": ("Admin & Office","General Admin","Safety Apparel"),
    "461818": ("Admin & Office","General Admin","Vision Protection"),
    "471300": ("Facilities / MRO","Janitorial","Cleaning & Janitorial Supplies"),
    "471399": ("Facilities / MRO","Janitorial","Salt & De-icing Products"),
    "500000": ("Food & Catering","Food & Beverage","Other Food & Beverage"),
    "501900": ("Food & Catering","Food & Beverage","Prepared Food"),
    "502000": ("Food & Catering","Food & Beverage","Beverages"),
    "510000": ("Clinical / Healthcare","Medical Supplies","Drug & Pharmaceutical"),
    "512100": ("Clinical / Healthcare","Medical Supplies","Drugs & Vaccines"),
    "512101": ("Clinical / Healthcare","Medical Supplies","Pharmacy Supplies"),
    "521000": ("Admin & Office","Furniture & Fixtures","Floor Covering & Carpet"),
    "521200": ("Admin & Office","Furniture & Fixtures","Bedclothes & Linens"),
    "521400": ("Admin & Office","Furniture & Fixtures","Domestic Appliances"),
    "521600": ("IT","IT Hardware & Peripherals","Consumer Electronics"),
    "531000": ("Admin & Office","General Admin","Clothing & Uniforms"),
    "531100": ("Admin & Office","General Admin","Footwear"),
    "560000": ("Admin & Office","Furniture & Fixtures","Other Furniture & Furnishings"),
    "561100": ("Admin & Office","Furniture & Fixtures","Commercial & Industrial Furniture"),
    "561200": ("Admin & Office","Furniture & Fixtures","Classroom & Institutional Furniture"),
    "600000": ("Admin & Office","Published / Books","Educational Materials & Supplies"),
    "601000": ("Admin & Office","Published / Books","Teaching Aids & Materials"),
    "601016": ("Admin & Office","Published / Books","Education Certificates & Diplomas"),
    "601100": ("Admin & Office","Published / Books","Classroom Supplies"),
    "721020": ("Facilities / MRO","MRO Supplies","Coating & Caulking"),
    "801416": ("Printing, Marketing & Communications","Promotional / Branded Items","Promotional Products & Giveaways"),
    "999088": ("Admin & Office","General Admin","Other Supplies & Equipment"),
    "999089": ("Facilities / MRO","MRO Supplies","Electrical Components & Supplies"),
    "999095": ("Admin & Office","Published / Books","Books & Subscriptions"),

    # ── SERVICES (SVCS) ───────────────────────────────────────
    "241100": ("Services","Consulting & Advisory","Storage Services"),
    "411300": ("Research & Laboratory","Scientific Services","Service Agreements - Lab & Scientific Equipment"),
    "551000": ("Printing, Marketing & Communications","Printing","Print Media"),
    "720000": ("Facilities / MRO","Trades Services","Building, Construction & Maintenance"),
    "721000": ("Facilities / MRO","Trades Services","Building Support, Maintenance & Repair"),
    "721013": ("Facilities / MRO","Trades Services","Elevator Maintenance"),
    "721014": ("Facilities / MRO","Trades Services","Locksmith"),
    "721015": ("Facilities / MRO","Trades Services","General Repair"),
    "721016": ("Facilities / MRO","Trades Services","Roofing & Siding"),
    "721017": ("Facilities / MRO","Trades Services","Concrete Work"),
    "721018": ("Facilities / MRO","Trades Services","Exterior Cleaning"),
    "721019": ("Facilities / MRO","Trades Services","Interior Finishing"),
    "721021": ("Facilities / MRO","Trades Services","Pest Control"),
    "721022": ("Facilities / MRO","Trades Services","Electrical Services"),
    "721023": ("Facilities / MRO","Trades Services","Plumbing, Heating & HVAC"),
    "721024": ("Facilities / MRO","Trades Services","Painting & Mold Remediation"),
    "721025": ("Facilities / MRO","Trades Services","Masonry & Stonework"),
    "721026": ("Facilities / MRO","Trades Services","Carpentry"),
    "721027": ("Facilities / MRO","Trades Services","Flooring Install & Repair"),
    "721028": ("Facilities / MRO","Trades Services","Refurbishing"),
    "721029": ("Facilities / MRO","Trades Services","Grounds Maintenance"),
    "761100": ("Facilities / MRO","Janitorial","Cleaning & Janitorial Services"),
    "761200": ("Facilities / MRO","Janitorial","Hazmat Waste Removal"),
    "761215": ("Facilities / MRO","Janitorial","Refuse Collection & Disposal"),
    "771000": ("Facilities / MRO","Trades Services","Environmental Services"),
    "781018": ("Travel, Events & Hospitality","Transportation","Relocation & Moving"),
    "781021": ("Travel, Events & Hospitality","Transportation","Freight & Shipping"),
    "781022": ("Travel, Events & Hospitality","Transportation","Small Parcel & Courier"),
    "781100": ("Travel, Events & Hospitality","Transportation","Passenger Transport"),
    "781116": ("Travel, Events & Hospitality","Transportation","Vehicle Rental"),
    "781117": ("Travel, Events & Hospitality","Transportation","Parking & Toll Fees"),
    "781118": ("Travel, Events & Hospitality","Transportation","Chartered Bus"),
    "781119": ("Travel, Events & Hospitality","Transportation","Taxi, Car & Shuttle"),
    "781200": ("Travel, Events & Hospitality","Transportation","Packing & Handling"),
    "781300": ("Travel, Events & Hospitality","Transportation","Storage"),
    "781801": ("Facilities / MRO","MRO Supplies","Vehicle Maintenance & Repair"),
    "801000": ("Services","Consulting & Advisory","Management Advisory & Support"),
    "801019": ("Services","Consulting & Advisory","Independent Contractor"),
    "801100": ("Services","Consulting & Advisory","Human Resources"),
    "801116": ("Services","Temp Staffing","Temporary Personnel"),
    "801117": ("Services","Temp Staffing","Recruiting & Executive Search"),
    "801200": ("Services","Legal","Legal Services"),
    "801300": ("Utilities & Occupancy","Rent / Lease / Occupancy","Real Estate"),
    "801400": ("Printing, Marketing & Communications","Marketing / Advertising","Marketing & Distribution"),
    "801417": ("Travel, Events & Hospitality","Event / Venue","Event Management Supplies"),
    "801600": ("Services","Consulting & Advisory","Business Administration"),
    "801616": ("Services","Consulting & Advisory","Business Facilities Oversight"),
    "801700": ("Services","Consulting & Advisory","Document & Records Storage"),
    "811000": ("Services","Consulting & Advisory","Engineering - Other"),
    "811015": ("Services","Consulting & Advisory","Civil Engineering"),
    "811016": ("Services","Consulting & Advisory","Mechanical Engineering"),
    "811017": ("Services","Consulting & Advisory","Electrical & Electronic Engineering"),
    "811027": ("Services","Consulting & Advisory","Architecture"),
    "811028": ("Services","Consulting & Advisory","Interior Design"),
    "811100": ("IT","IT Services","Computer/IT Consulting"),
    "811115": ("IT","IT Services","Engineering - Software or Hardware"),
    "811116": ("IT","IT Services","Computer Programmers"),
    "811117": ("IT","IT Services","Management Information Systems"),
    "811118": ("IT","IT Services","System Administrators"),
    "811119": ("IT","IT Services","Information Retrieval Systems"),
    "811120": ("IT","IT Services","Data Services"),
    "811121": ("IT","IT Services","Internet Services"),
    "811122": ("IT","IT Services","Software Maintenance & Support"),
    "811123": ("IT","IT Services","Hardware Maintenance & Support"),
    "811607": ("IT","Network / Telecom","Telecommunications Services"),
    "811617": ("IT","Network / Telecom","Telecom Equipment Installation & Maintenance"),
    "812000": ("IT","IT Services","Photocopier Maintenance & Support"),
    "821000": ("Printing, Marketing & Communications","Marketing / Advertising","Advertising & Media"),
    "821016": ("Printing, Marketing & Communications","Marketing / Advertising","Broadcast Media"),
    "821100": ("Printing, Marketing & Communications","Marketing / Advertising","Writing & Translations"),
    "821120": ("Printing, Marketing & Communications","Marketing / Advertising","Language & Interpretation"),
    "821200": ("Printing, Marketing & Communications","Printing","Reproduction Services"),
    "821218": ("Printing, Marketing & Communications","Printing","Printing, Binding & Publishing"),
    "821300": ("Printing, Marketing & Communications","Marketing / Advertising","Photographic Services"),
    "821400": ("Printing, Marketing & Communications","Marketing / Advertising","Graphic & Website Design"),
    "821500": ("Travel, Events & Hospitality","Event / Venue","Professional Artists & Performers"),
    "831000": ("Utilities & Occupancy","Electric / Gas","Utility Services"),
    "831100": ("IT","Network / Telecom","Telecommunications Media"),
    "831118": ("IT","Network / Telecom","Television & Cable"),
    "840000": ("Services","Audit / Compliance","Financial or Insurance Services"),
    "841100": ("Services","Audit / Compliance","Accounting & Auditing"),
    "841116": ("Services","Audit / Compliance","Audit Services"),
    "841200": ("Services","Consulting & Advisory","Banking & Investment"),
    "841300": ("Services","Consulting & Advisory","Employee Insurance & Retirement"),
    "851220": ("Clinical / Healthcare","Clinical Services","Dental Services"),
    "851300": ("Research & Laboratory","Scientific Services","Medical Science Research & Experiment"),
    "851500": ("Food & Catering","Food & Beverage","Food & Nutrition Services"),
    "851600": ("Clinical / Healthcare","Medical Equipment","Medical Equipment Maintenance"),
    "861200": ("Services","Consulting & Advisory","Educational Institution Services"),
    "861600": ("Services","Consulting & Advisory","Student Recruiting"),
    "901100": ("Travel, Events & Hospitality","Lodging","Hotels & Lodging & Meeting/Event Facilities"),
    "901200": ("Travel, Events & Hospitality","Transportation","Travel Facilitation"),
    "901300": ("Travel, Events & Hospitality","Event / Venue","Performing Arts"),
    "901500": ("Travel, Events & Hospitality","Event / Venue","Entertainment"),
    "921000": ("Services","Consulting & Advisory","Public Safety"),
    "921214": ("Services","Consulting & Advisory","Campus Security"),
    "921217": ("IT","IT Services","IT Security Services"),
    "931500": ("Inter-Entity / Transfers","City/State/Federal","Public Administration & Finance"),
    "999052": ("Services","Consulting & Advisory","Training Services"),
    "999055": ("IT","IT Services","Technology - Other"),
    "999058": ("Services","Consulting & Advisory","Teaching & Instructional Services"),
    "999065": ("Travel, Events & Hospitality","Transportation","Professional Development & Continuing Ed"),
    "999066": ("Printing, Marketing & Communications","Marketing / Advertising","Print Advertising"),
    "999067": ("Printing, Marketing & Communications","Marketing / Advertising","Photography & Videography"),
    "999068": ("Clinical / Healthcare","Medical Supplies","Pharmacy Services"),
    "999069": ("IT","IT Services","Other IT-Related Services"),
    "999070": ("Services","Consulting & Advisory","Other Services"),
    "999073": ("Clinical / Healthcare","Clinical Services","Medical Professional Services"),
    "999074": ("Clinical / Healthcare","Medical Equipment","Medical Equipment Maintenance & Refurbishment"),
    "999076": ("Research & Laboratory","Scientific Services","Laboratory Services"),
    "999078": ("Facilities / MRO","Janitorial","Industrial & Specialized Laundry"),
    "999080": ("Facilities / MRO","MRO Supplies","General Equipment Maintenance & Refurbishment"),
    "999083": ("Services","Consulting & Advisory","Data Analysis"),
    "999085": ("Services","Consulting & Advisory","Consulting Services"),
    "999087": ("IT","IT Services","A/V Installation & Maintenance"),
    "999096": ("Inter-Entity / Transfers","Direct Pay","Direct Pay"),
    "999097": ("Admin & Office","Published / Books","Library Services"),
}

# ══════════════════════════════════════════════════════════════
# PASS 2 — VENDOR ALWAYS-LIST
# ══════════════════════════════════════════════════════════════
# Vendors that should always override their commodity code (Pass 0)
# Use this for cases where the commodity code gives a misleading bucket
VENDOR_OVERRIDES = {
    # ── Original overrides ────────────────────────────────────────────────────
    "aramark":          ("Food & Catering","Food & Beverage","Dining Services","vendor:aramark"),
    "barnes & noble":   ("Admin & Office","Published / Books","Textbooks & Course Materials","vendor:barnes_noble"),
    "exela enterprise": ("Admin & Office","Office Supplies","Postage & Mail Services","vendor:exela"),
    "dbhds":            ("Inter-Entity / Transfers","City/State/Federal","State Behavioral Health","vendor:dbhds"),

    # ── Utilities & Occupancy ─────────────────────────────────────────────────
    # Telecom providers misrouted to IT via commodity code
    "verizon":          ("Utilities & Occupancy","Electric / Gas","Telecom","vendor:verizon"),
    "comcast":          ("Utilities & Occupancy","Electric / Gas","Telecom","vendor:comcast"),
    # Water cooler / bottled water service misrouted to Research or Admin
    "primo water":      ("Utilities & Occupancy","Electric / Gas","Utilities","vendor:primo_water"),
    # Electrical contractor flagged in Travel
    "capital electric": ("Utilities & Occupancy","Electric / Gas","Electricity","vendor:capital_electric"),

    # ── Facilities / MRO ──────────────────────────────────────────────────────
    # Waste hauler misrouted to Travel via GL account
    "county waste":                  ("Facilities / MRO","Trades Services","General MRO","vendor:county_waste"),
    # Disaster restoration company misrouted to Prof Services and IT
    "belfor":                        ("Facilities / MRO","Trades Services","General Contractor","vendor:belfor"),
    # Fire & security contractor misrouted to IT
    "vsc fire":                      ("Facilities / MRO","Trades Services","General MRO","vendor:vsc_fire"),
    # Electrical supply houses misrouted to Admin or IT
    "electrical equipment co":       ("Facilities / MRO","MRO Supplies","General MRO","vendor:electrical_equip"),
    "old dominion electrical":       ("Facilities / MRO","MRO Supplies","General MRO","vendor:od_electrical"),
    "modern electrical supplies":    ("Facilities / MRO","MRO Supplies","General MRO","vendor:modern_elec"),
    "maurice electrical supply":     ("Facilities / MRO","MRO Supplies","General MRO","vendor:maurice_elec"),
    # Waste treatment (Qatar campus) misrouted to Prof Services
    "al haya waste":                 ("Facilities / MRO","Trades Services","General MRO","vendor:al_haya_waste"),

    # ── Professional Services ─────────────────────────────────────────────────
    # Engineering consultancy misrouted to Facilities via commodity
    "schnabel engineering":          ("Capital Projects & Construction","Architecture & Engineering","Engineering","vendor:schnabel"),
    # Insurance brokers/carriers misrouted to Inter-Entity or Travel
    "mcgriff insurance":             ("Services","Consulting & Advisory","Consulting","vendor:mcgriff"),
    "guardian life insurance":       ("Services","Consulting & Advisory","Consulting","vendor:guardian_life"),
    "lincoln national life":         ("Services","Consulting & Advisory","Consulting","vendor:lincoln_national"),
    # Audit firm misrouted to Travel
    "deloitte and touche":           ("Services","Consulting & Advisory","Consulting","vendor:deloitte"),
    # Professional associations misrouted to Travel or Admin
    "virginia academy of science engineering and medicine": ("Services","Consulting & Advisory","Consulting","vendor:vasem"),
    "american society for engineering education":           ("Services","Consulting & Advisory","Consulting","vendor:asee"),
    "society of automotive engineering":                    ("Services","Consulting & Advisory","Consulting","vendor:sae"),

    # ── Clinical / Healthcare ─────────────────────────────────────────────────
    # Medical device company misrouted to Prof Services and Research
    "philips healthcare":            ("Clinical / Healthcare","Medical Equipment","Medical Devices","vendor:philips_hc"),
    # Hospitals misrouted to Prof Services, Travel, or Uncategorized
    "rhode island hospital":         ("Clinical / Healthcare","Patient/Participant Payments","Stipends","vendor:ri_hospital"),
    "massachusetts general hospital":("Clinical / Healthcare","Patient/Participant Payments","Stipends","vendor:mgh"),
    "chesapeake general hospital":   ("Clinical / Healthcare","Patient/Participant Payments","Stipends","vendor:chesapeake_gen"),
    "mary washington hospital":      ("Clinical / Healthcare","Patient/Participant Payments","Stipends","vendor:mary_washington"),
    "sentara martha jefferson":      ("Clinical / Healthcare","Patient/Participant Payments","Stipends","vendor:sentara_mj"),
    "bon secours richmond community":("Clinical / Healthcare","Patient/Participant Payments","Stipends","vendor:bon_secours"),
    "sibley memorial hospital":      ("Clinical / Healthcare","Patient/Participant Payments","Stipends","vendor:sibley"),
    "cleveland clinic":              ("Clinical / Healthcare","Patient/Participant Payments","Stipends","vendor:cleveland_clinic"),
    "st. jude children":             ("Clinical / Healthcare","Patient/Participant Payments","Stipends","vendor:st_jude"),
    "central state hospital":        ("Clinical / Healthcare","Patient/Participant Payments","Stipends","vendor:central_state_hosp"),
    "encompass health rehabilitation":("Clinical / Healthcare","Patient/Participant Payments","Stipends","vendor:encompass_health"),
    # Hospital association misrouted to IT and Prof Services
    "american hospital association": ("Clinical / Healthcare","Patient/Participant Payments","Stipends","vendor:aha"),
    # Pharmacy and clinical orgs misrouted to Research or Admin
    "wedgewood village pharmacy":    ("Clinical / Healthcare","Medical Supplies","Clinical Supplies","vendor:wedgewood_pharm"),
    "national association of boards of pharmacy": ("Clinical / Healthcare","Medical Supplies","Clinical Supplies","vendor:nabp"),
    "pediatric pharmacy advocacy":   ("Clinical / Healthcare","Patient/Participant Payments","Stipends","vendor:ppag"),
    "student college of clinical pharmacy": ("Clinical / Healthcare","Patient/Participant Payments","Stipends","vendor:sccp"),

    # Pharma manufacturer confirmed Clinical (previously skipped, now resolved)
    "takeda pharmaceuticals":        ("Clinical / Healthcare","Medical Supplies","Clinical Supplies","vendor:takeda"),

    # ── Admin & Office ────────────────────────────────────────────────────────
    # Xerox misrouted to Utilities and Printing via commodity code
    "xerox":                         ("Admin & Office","Office Supplies","Postage & Mail Services","vendor:xerox"),
    # Paper/print supply company misrouted to Prof Services
    "shepherd specialty paper":      ("Admin & Office","Office Supplies","Postage & Mail Services","vendor:shepherd_paper"),

    # ── Travel, Events & Hospitality — Hotels ────────────────────────────────
    # Doha hotels hitting Food & Catering or Admin (Qatar campus spend)
    "grand hyatt doha":              ("Travel, Events & Hospitality","Lodging","Hotel","vendor:grand_hyatt_doha"),
    "park hyatt doha":               ("Travel, Events & Hospitality","Lodging","Hotel","vendor:park_hyatt_doha"),
    "intercontinental hotel doha":   ("Travel, Events & Hospitality","Lodging","Hotel","vendor:ic_doha"),
    "la cigale hotel":               ("Travel, Events & Hospitality","Lodging","Hotel","vendor:la_cigale"),
    "plaza doha hotel":              ("Travel, Events & Hospitality","Lodging","Hotel","vendor:plaza_doha"),
    "sheraton grand doha":           ("Travel, Events & Hospitality","Lodging","Hotel","vendor:sheraton_doha"),
    "best plaza west bay":           ("Travel, Events & Hospitality","Lodging","Hotel","vendor:best_plaza"),
    "barahat msheireb":              ("Travel, Events & Hospitality","Lodging","Hotel","vendor:barahat"),
    "delta hotels city center doha": ("Travel, Events & Hospitality","Lodging","Hotel","vendor:delta_doha"),
    # Domestic hotels hitting Food & Catering or Facilities
    "westin hotel":                  ("Travel, Events & Hospitality","Lodging","Hotel","vendor:westin"),
    "omni charlottesville":          ("Travel, Events & Hospitality","Lodging","Hotel","vendor:omni_cville"),
    "omni richmond":                 ("Travel, Events & Hospitality","Lodging","Hotel","vendor:omni_richmond"),
    "hotel roanoke":                 ("Travel, Events & Hospitality","Lodging","Hotel","vendor:hotel_roanoke"),
    "hyatt house richmond":          ("Travel, Events & Hospitality","Lodging","Hotel","vendor:hyatt_richmond"),

    # ── Travel, Events & Hospitality — Events & Entertainment ────────────────
    # Entertainment/event vendors misrouted to Admin or Prof Services
    "harlem globetrotters":          ("Travel, Events & Hospitality","Event / Venue","Entertainment","vendor:harlem_globetrotters"),
    "event technologies inc":        ("Travel, Events & Hospitality","Event / Venue","AV Services","vendor:event_tech"),
    "lightning event":               ("Travel, Events & Hospitality","Event / Venue","Events","vendor:lightning_event"),
    "prime time party":              ("Travel, Events & Hospitality","Event / Venue","Events","vendor:prime_time_party"),
    "commonwealth event company":    ("Travel, Events & Hospitality","Event / Venue","Events","vendor:commonwealth_event"),
}

VENDOR_MAP = {
    # Utilities
    "dominion virginia power":       ("Utilities & Occupancy","Electric / Gas","Electricity","vendor:dominion"),
    "dominion energy":               ("Utilities & Occupancy","Electric / Gas","Electricity","vendor:dominion"),
    "atrium campus":                 ("Utilities & Occupancy","Rent / Lease / Occupancy","Space Lease","vendor:atrium"),
    # Research & Lab
    "fisher scientific":             ("Research & Laboratory","Lab Consumables","Lab Supplies","vendor:fisher"),
    "thermo fisher":                 ("Research & Laboratory","Lab Equipment & Instruments","Instruments","vendor:thermofisher"),
    "leica microsystems":            ("Research & Laboratory","Lab Equipment & Instruments","Microscopy","vendor:leica"),
    "beckman coulter":               ("Research & Laboratory","Lab Equipment & Instruments","Instruments","vendor:beckman"),
    "meso scale discovery":          ("Research & Laboratory","Lab Consumables","Assay Kits","vendor:msd"),
    "eckert & ziegler":              ("Research & Laboratory","Scientific Services","Radiopharmaceutical","vendor:eckert"),
    "bruker":                        ("Research & Laboratory","Lab Equipment & Instruments","NMR/MS Instruments","vendor:bruker"),
    "revvity":                       ("Research & Laboratory","Lab Equipment & Instruments","Life Science Instruments","vendor:revvity"),
    "perkinelmer":                   ("Research & Laboratory","Lab Equipment & Instruments","Instruments","vendor:perkinelmer"),
    "allentown":                     ("Research & Laboratory","Lab Equipment & Instruments","Animal Housing","vendor:allentown"),
    "sigma-aldrich":                 ("Research & Laboratory","Chemicals & Reagents","Reagents","vendor:sigma"),
    "vwr international":             ("Research & Laboratory","Lab Consumables","Lab Supplies","vendor:vwr"),
    "bio-rad":                       ("Research & Laboratory","Lab Consumables","Reagents & Kits","vendor:biorad"),
    "qiagen":                        ("Research & Laboratory","Lab Consumables","Molecular Biology","vendor:qiagen"),
    # Clinical
    "fujifilm sonosite":             ("Clinical / Healthcare","Medical Equipment","Ultrasound","vendor:fujifilm"),
    "medline":                       ("Clinical / Healthcare","Medical Supplies","Clinical Supplies","vendor:medline"),
    "cardinal health":               ("Clinical / Healthcare","Medical Supplies","Clinical Supplies","vendor:cardinal"),
    "mckesson":                      ("Clinical / Healthcare","Medical Supplies","Pharmaceutical","vendor:mckesson"),
    # IT
    "shi international":             ("IT","IT Hardware & Peripherals","IT Reseller","vendor:shi"),
    "cdw":                           ("IT","IT Hardware & Peripherals","IT Reseller","vendor:cdw"),
    "dell":                          ("IT","IT Hardware & Peripherals","Computers","vendor:dell"),
    "apple":                         ("IT","IT Hardware & Peripherals","Computers","vendor:apple"),
    "microsoft":                     ("IT","IT Software / SaaS","Microsoft Licenses","vendor:microsoft"),
    "adobe":                         ("IT","IT Software / SaaS","Adobe Licenses","vendor:adobe"),
    "zoom":                          ("IT","IT Software / SaaS","Video Conferencing","vendor:zoom"),
    "collaborative technologies":    ("IT","IT Services","AV & Communications","vendor:collab_tech"),
    "cambridge computer":            ("IT","IT Hardware & Peripherals","IT Reseller","vendor:cambridge_comp"),
    # Facilities / MRO
    "w w grainger":                  ("Facilities / MRO","MRO Supplies","General MRO","vendor:grainger"),
    "grainger":                      ("Facilities / MRO","MRO Supplies","General MRO","vendor:grainger"),
    "fastenal":                      ("Facilities / MRO","MRO Supplies","Fasteners/MRO","vendor:fastenal"),
    "colonial webb":                 ("Facilities / MRO","Trades Services","HVAC & Mechanical","vendor:colonial_webb"),
    "glave & holmes":                ("Capital Projects & Construction","Architecture & Engineering","Architecture","vendor:glave"),
    "landscape workshop":            ("Facilities / MRO","Trades Services","Grounds Maintenance","vendor:landscape"),
    "timmons group":                 ("Capital Projects & Construction","Architecture & Engineering","Engineering/GIS","vendor:timmons"),
    # Professional Services
    "practicewise":                  ("Services","Consulting & Advisory","Healthcare Analytics","vendor:practicewise"),
    "economic modeling":             ("Services","Consulting & Advisory","Economic Research","vendor:emsi"),
    # Travel & Events
    "refresh music group":           ("Travel, Events & Hospitality","Event / Venue","Entertainment","vendor:refresh_music"),
    "van evera":                     ("Travel, Events & Hospitality","Event / Venue","AV Services","vendor:van_evera"),
    "marriott":                      ("Travel, Events & Hospitality","Lodging","Hotel","vendor:marriott"),
    "hilton":                        ("Travel, Events & Hospitality","Lodging","Hotel","vendor:hilton"),
    # Capital Construction & A&E
    "barton malow":                  ("Capital Projects & Construction","General Contractor / CM","General Contractor","vendor:barton_malow"),
    "hanbury evans":                 ("Capital Projects & Construction","Architecture & Engineering","Architecture / A&E","vendor:hanbury_evans"),
    "woodland construction":         ("Capital Projects & Construction","General Contractor / CM","General Contractor","vendor:woodland_const"),
    "smithgroup":                    ("Capital Projects & Construction","Architecture & Engineering","Architecture / A&E","vendor:smithgroup"),

    # ── IT — AV / UC / Low-Voltage Integrators ───────────────────────────────
    "avi-spl":                       ("IT","IT Services","AV & Media Services","vendor:avispl"),
    "avispl":                        ("IT","IT Services","AV & Media Services","vendor:avispl"),
    "active technology solutions":   ("IT","IT Services","AV & Media Services","vendor:active_tech"),
    "virginia integrated communication": ("IT","IT Services","AV & Media Services","vendor:vicom"),
    "vicom":                         ("IT","IT Services","AV & Media Services","vendor:vicom"),

    # ── IT — Software / SaaS ─────────────────────────────────────────────────
    "carahsoft":                     ("IT","IT Software / SaaS","Software","vendor:carahsoft"),
    "givzey":                        ("IT","IT Software / SaaS","Software","vendor:givzey"),
    "hudl":                          ("IT","IT Software / SaaS","Software","vendor:hudl"),
    "ideagen":                       ("IT","IT Software / SaaS","Software","vendor:ideagen"),
    "instructure":                   ("IT","IT Software / SaaS","Software","vendor:instructure"),
    "leepfrog":                      ("IT","IT Software / SaaS","Software","vendor:leepfrog"),
    "cbord group":                   ("IT","IT Software / SaaS","Software","vendor:cbord"),

    # ── IT — Hardware / Resellers / Systems ──────────────────────────────────
    "cas severn":                    ("IT","IT Hardware & Peripherals","IT Reseller","vendor:cas_severn"),
    "ip datasystems":                ("IT","IT Hardware & Peripherals","IT Reseller","vendor:ip_datasystems"),
    "colorid":                       ("IT","IT Hardware & Peripherals","IT Equipment","vendor:colorid"),

    # ── Research & Laboratory ─────────────────────────────────────────────────
    "andelyn biosciences":           ("Research & Laboratory","Scientific Services","Core Facilities","vendor:andelyn"),
    "charles river laboratories":    ("Research & Laboratory","Scientific Services","Core Facilities","vendor:charles_river"),
    "r&d systems":                   ("Research & Laboratory","Chemicals & Reagents","Reagents & Chemicals","vendor:rd_systems"),
    "bio-techne":                    ("Research & Laboratory","Chemicals & Reagents","Reagents & Chemicals","vendor:biotechne"),
    "jackson laboratory":            ("Research & Laboratory","Lab Consumables","Lab Supplies","vendor:jackson_lab"),
    "transnetyx":                    ("Research & Laboratory","Scientific Services","Core Facilities","vendor:transnetyx"),

    # ── Clinical / Healthcare — Dental ───────────────────────────────────────
    "bien-air":                      ("Clinical / Healthcare","Medical Equipment","Medical Devices","vendor:bien_air"),
    "dentsply sirona":               ("Clinical / Healthcare","Medical Equipment","Medical Devices","vendor:dentsply"),
    "henry schein":                  ("Clinical / Healthcare","Medical Supplies","Clinical Supplies","vendor:henry_schein"),
    "ivoclar vivadent":              ("Clinical / Healthcare","Medical Equipment","Medical Devices","vendor:ivoclar"),

    # ── Facilities / MRO ─────────────────────────────────────────────────────
    "clean harbors":                 ("Facilities / MRO","Trades Services","General MRO","vendor:clean_harbors"),
    "priority elevator":             ("Facilities / MRO","Trades Services","General MRO","vendor:priority_elevator"),
    "creative office environments":  ("Facilities / MRO","MRO Supplies","General MRO","vendor:creative_office_env"),
    "pmc commercial interiors":      ("Facilities / MRO","MRO Supplies","General MRO","vendor:pmc_interiors"),

    # ── Printing, Marketing & Communications ─────────────────────────────────
    "swish llc":                     ("Printing, Marketing & Communications","Printing","Print Services","vendor:swish"),
    "tk promotions":                 ("Printing, Marketing & Communications","Printing","Print Services","vendor:tk_promotions"),
    "suburban remodeling":           ("Facilities / MRO","Trades Services","General Contractor","vendor:suburban_remod"),
    "arw contracting":               ("Facilities / MRO","Trades Services","General Contractor","vendor:arw_contracting"),
    "stoker construction":           ("Facilities / MRO","Trades Services","General Contractor","vendor:stoker"),
    "baskervill":                    ("Capital Projects & Construction","Architecture & Engineering","Architecture / A&E","vendor:baskervill"),
    "raymond engineering":           ("Capital Projects & Construction","Architecture & Engineering","Engineering","vendor:raymond_eng"),
    "rrmm architects":               ("Capital Projects & Construction","Architecture & Engineering","Architecture / A&E","vendor:rrmm"),
    "morgan keller":                 ("Facilities / MRO","Trades Services","General Contractor","vendor:morgan_keller"),
    "dunbar pllc":                   ("Capital Projects & Construction","Architecture & Engineering","Architecture / A&E","vendor:dunbar"),
    "marshall craft":                ("Capital Projects & Construction","Architecture & Engineering","Architecture / A&E","vendor:marshall_craft"),
    "facility dynamics":             ("Capital Projects & Construction","Architecture & Engineering","Engineering","vendor:facility_dynamics"),
    "kimley-horn":                   ("Capital Projects & Construction","Architecture & Engineering","Engineering","vendor:kimley_horn"),
    "montgomery consulting":         ("Facilities / MRO","Trades Services","General Contractor","vendor:montgomery_const"),
    "ayers saint gross":             ("Capital Projects & Construction","Architecture & Engineering","Architecture / A&E","vendor:ayers_saint_gross"),
    "tate engineering":              ("Facilities / MRO","Trades Services","HVAC & Mechanical","vendor:tate_eng"),
    "affiliated engineers":          ("Capital Projects & Construction","Architecture & Engineering","Engineering","vendor:affiliated_eng"),
    "davis and green":               ("Facilities / MRO","Trades Services","General Contractor","vendor:davis_green"),
    "trc engineers":                 ("Capital Projects & Construction","Architecture & Engineering","Engineering","vendor:trc_engineers"),
    # Specific vendor overrides (commodity code gives wrong bucket)
    "aramark":                       ("Food & Catering","Food & Beverage","Dining Services","vendor:aramark"),
    "barnes & noble":                ("Admin & Office","Published / Books","Textbooks & Course Materials","vendor:barnes_noble"),
    "exela enterprise":              ("Admin & Office","Office Supplies","Postage & Mail Services","vendor:exela"),
    "trane":                         ("Facilities / MRO","Trades Services","HVAC & Mechanical","vendor:trane"),
    "pmc commercial interiors":      ("Admin & Office","Furniture & Fixtures","Office Furniture","vendor:pmc_interiors"),
    "sycom technologies":            ("IT","Network / Telecom","Network Infrastructure","vendor:sycom"),
    "epitome networks":              ("IT","IT Hardware & Peripherals","AV & Network Equipment","vendor:epitome"),
    "rtw media":                     ("IT","IT Services","AV & Media Services","vendor:rtw_media"),
    "dbhds":                         ("Inter-Entity / Transfers","City/State/Federal","State Behavioral Health","vendor:dbhds"),
    # Inter-Entity
    "candex":                        ("Inter-Entity / Transfers","Candex / Pass-through","Payment Rail","vendor:candex"),
    "vcu health system":             ("Inter-Entity / Transfers","Internal Health System","VCU Health","vendor:vcuhs"),
    "virginia commonwealth univer":  ("Inter-Entity / Transfers","Internal VCU Entity","VCU Entity","vendor:vcu_entity"),
    "boston university":             ("Inter-Entity / Transfers","Other Universities","University","vendor:boston_u"),
    "west virginia university":      ("Inter-Entity / Transfers","Other Universities","University","vendor:wvu"),
    "virginia department of health": ("Inter-Entity / Transfers","City/State/Federal","State Agency","vendor:vdh"),
    "city of richmond":              ("Inter-Entity / Transfers","City/State/Federal","Local Government","vendor:richmond"),
    "us dhhs":                       ("Inter-Entity / Transfers","City/State/Federal","Federal Agency","vendor:dhhs"),
    "treasurer of virginia":         ("Inter-Entity / Transfers","City/State/Federal","State","vendor:va_treasurer"),
    "commonwealth of virginia":      ("Inter-Entity / Transfers","City/State/Federal","State","vendor:va_state"),
}

# ══════════════════════════════════════════════════════════════
# PASS 3 — CATEGORY LEVEL 1 METADATA
# ══════════════════════════════════════════════════════════════
CATEGORY_L1_MAP = {
    "laboratory and measuring":     ("Research & Laboratory","Lab Equipment & Instruments","Scientific Instruments"),
    "chemicals including":          ("Research & Laboratory","Chemicals & Reagents","Reagents & Chemicals"),
    "medical equipment":            ("Clinical / Healthcare","Medical Equipment","Medical Devices"),
    "medical supplies":             ("Clinical / Healthcare","Medical Supplies","Clinical Supplies"),
    "information technology":       ("IT","IT Hardware & Peripherals","IT Equipment"),
    "telecommunications":           ("IT","Network / Telecom","Telecom"),
    "power generation":             ("Utilities & Occupancy","Electric / Gas","Power"),
    "published products":           ("Admin & Office","Published / Books","Books & Subscriptions"),
    "furniture and furnishings":    ("Admin & Office","Furniture & Fixtures","Furniture"),
    "food beverage":                ("Food & Catering","Food & Beverage","Food"),
    "manufacturing components":     ("Facilities / MRO","MRO Supplies","Manufacturing/MRO"),
    "building and construction":    ("Facilities / MRO","Trades Services","Construction"),
    "cleaning":                     ("Facilities / MRO","Janitorial","Cleaning"),
    "printing":                     ("Printing, Marketing & Communications","Printing","Print Services"),
    "travel":                       ("Travel, Events & Hospitality","Travel","Transportation"),
}

# ══════════════════════════════════════════════════════════════
# PASS 4 — KEYWORD / REGEX SCAN
# ══════════════════════════════════════════════════════════════
KEYWORD_PATTERNS = [
    ("IT","IT Hardware & Peripherals","Computers",
     r"\b(laptop|notebook|desktop|workstation|monitor|docking|keyboard|mouse|tablet|ipad|iphone|chromebook|server|storage|nas|san|switch|router|firewall|access.?point|wifi|cisco|meraki|juniper|aruba|ups|uninterruptible|scanner)\b",
     "kw:it_hardware"),
    ("IT","IT Software / SaaS","Software",
     r"\b(licen[cs]e|subscription|saas|software|renewal|maintenance.?agree|cloud.?hosting|vmware|office.?365|azure|aws|oracle|salesforce|zoom|slack|servicenow|jira|atlassian|matlab|stata|spss|endnote|adobe)\b",
     "kw:it_software"),
    ("Research & Laboratory","Lab Consumables","Lab Supplies",
     r"\b(reagent|assay|antibody|enzyme|buffer|pcr|qpcr|elisa|hplc|gcms|nmr|centrifuge|incubator|cryogenic|liquid nitrogen|cell culture|dmem|fbs|serum|pipett|microplate|vial|flask|beaker|glassware|eppendorf|cuvette)\b",
     "kw:lab_consumables"),
    ("Research & Laboratory","Lab Equipment & Instruments","Instruments",
     r"\b(mass spec|flow cyt|sequenc|imaging system|analyzer|electrophoresis|western blot|thermocycler|autoclave|biosafety cabinet|fume hood|lyophiliz|freeze.?dry|spectrophotom|microscop)\b",
     "kw:lab_equipment"),
    ("Clinical / Healthcare","Patient/Participant Payments","Stipends",
     r"\b(stipend|participant payment|patient stipend|honorarium|honoraria|subject payment)\b",
     "kw:stipend"),
    # ── Capital Projects & Construction ───────────────────────────────────────
    ("Capital Projects & Construction","General Contractor / CM","Construction",
     r"\b(general contractor|\bGC\b|construction manager|\bCM\b|construction contract|renovation contract|new construction|building construction|sitework|earthwork|demolition|abatement|grading|foundation|concrete work|masonry|roofing contract|waterproof(ing)?|commissioning service|construction service|build-out|fit.?out|tenant improvement|\bTI\b\s+work|capital improvement|capital project|capital renewal|infrastructure project|deferred maintenance project)\b",
     "kw:construction"),
    ("Capital Projects & Construction","Architecture & Engineering","A&E Services",
     r"\b(architectural service|architecture service|engineering service|design service|design.?build|structural engineer|civil engineer|mep engineer|schematic design|design development|construction document|bid document|specification writing|commissioning engineer|owner.?s rep|\bCxA\b|project architect)\b",
     "kw:ae_services"),
    ("Capital Projects & Construction","Capital Equipment Installation","Equipment Installation",
     r"\b(equipment install|capital equipment install|major equipment|capital install|fume hood install|cleanroom|vivarium|lab build.?out|specialized equipment)\b",
     "kw:cap_equip_install"),
    ("Facilities / MRO","Trades Services","HVAC",
     r"\b(hvac|plumb(ing)?|electrical contractor|conduit|breaker|panel|service call|boiler|chiller|duct|compressor|generator|cooling tower)\b",
     "kw:trades"),
    ("Facilities / MRO","Janitorial","Cleaning",
     r"\b(janitorial|custodial|clean(ing|er)|mop|disinfect|sanitiz|floor wax|restroom supply)\b",
     "kw:janitorial"),
    ("Facilities / MRO","MRO Supplies","General MRO",
     r"\b(fastener|bolt|nut|screw|hinge|paint|caulk|seal|gasket|bearing|belt|pump|valve|fitting|coupling|bracket|mounting hardware)\b",
     "kw:mro"),
    ("Travel, Events & Hospitality","Lodging","Hotel",
     r"\b(lodging|hotel|inn|suites|accommodation|airbnb)\b",
     "kw:lodging"),
    ("Travel, Events & Hospitality","Transportation","Airfare",
     r"\b(airfare|flight|airline|delta|united|american airlines|southwest|amtrak|uber|lyft|per diem|mileage|car rental|rental car)\b",
     "kw:transportation"),
    ("Travel, Events & Hospitality","Event / Venue","Events",
     r"\b(conference|registration fee|venue|event space|banquet|gala|reception|award ceremony|sponsorship)\b",
     "kw:events"),
    # ── Services L2 taxonomy — keyword precedence order ─────────────────────
    # Rule: more-specific patterns fire first; generic fallback at end.
    # IT-delivered services (managed services, pro services, implementation)
    ("Services","IT Services","Managed IT / Support",
     r"\b(managed service|help.?desk|service.?desk|noc|network operation|it support|desktop support|end.?user support|break.?fix|remote monitoring|msp\b|it management)\b",
     "kw:svc_it_managed"),
    ("Services","IT Services","Implementation & Integration",
     r"\b(implementation|system integration|erp.?implementation|go.?live|configuration service|data migration|deployment service|system.?deployment|cloud migration|cutover)\b",
     "kw:svc_it_impl"),
    ("Services","IT Services","Cybersecurity",
     r"(?i)\b(cybersec|penetration.?test|pen.?test|vuln(erability)?.?(scan|assess)|soc\s+service|threat.?intel|siem.?service|incident.?response|security.?(audit|assess|review|program))\w*\b",
     "kw:svc_it_cyber"),
    # Legal
    ("Services","Legal Services","Legal Counsel",
     r"\b(attorney|legal counsel|law firm|legal fee|litigation|outside counsel|legal service|contract review|arbitrat|mediati|regulatory counsel|compliance counsel)\b",
     "kw:svc_legal"),
    # Staffing (before generic consulting)
    ("Services","Staffing & Temp Labor","Temp/Contract Staff",
     r"\b(staffing|temp(orary)?\s+staff|contingent\s+worker|contract\s+employee|contractor\s+placement|labor\s+hire|workforce\s+solution|contract\s+labor|contract\s+staff|temp\s+worker|agency\s+worker)\b",
     "kw:svc_staffing"),
    ("Services","Staffing & Temp Labor","Clinical Temp Staff",
     r"\b(travel nurse|locum|clinical staff|per diem nurse|nursing agenc|allied health staffing|clinical temp|healthcare staffing)\b",
     "kw:svc_staffing_clinical"),
    # Marketing & Creative
    ("Services","Marketing & Creative","Marketing / PR",
     r"\b(marketing\s+service|public\s+relation|pr\s+service|media\s+relation|branding\s+service|content\s+strateg|social\s+media|digital\s+market|seo\s+service|email\s+campaign|media\s+buy|advertising\s+service|creative\s+service|graphic\s+design|video\s+produc)\w*",
     "kw:svc_marketing"),
    # Training & Conferences
    ("Services","Training & Education","Professional Training",
     r"\b(training service|professional development|workforce training|e.?learning content|lms content|courseware|curriculum develop|instructional design|learning.?develop|certification program|compliance training)\b",
     "kw:svc_training"),
    # Research & Academic Services
    ("Services","Research & Academic Services","Academic / Research Support",
     r"\b(research service|academic service|scholarly service|subaward service|editorial service|peer review service|publication service|core facilit|bioinformatic service|genomic service|proteomics service|statistical consult)\b",
     "kw:svc_research"),
    # Clinical / Health Services (service delivery, not supplies)
    ("Services","Clinical Services","Clinical Service Delivery",
     r"\b(clinical service|health service|medical service|nursing service|therapy service|behavioral health service|telehealth|clinical trial service|patient care service|health program)\b",
     "kw:svc_clinical"),
    # Facilities-management services (not construction, not MRO supplies)
    ("Services","Facilities Services","Facilities Management",
     r"\b(facilities management|building management|property management|pest control|waste management|grounds service|landscape service|elevator service|fire inspection service|security guard|security service|parking management)\b",
     "kw:svc_facilities"),
    # Enrollment / OPM / Student services — high-value at universities
    ("Services","Consulting & Advisory","Enrollment / OPM",
     r"\b(enrollment manage|opm\b|online program manage|student recruit(ment)?\s+service|admission service|financial aid consult|retention service|student success service)\b",
     "kw:svc_enrollment"),
    # Other Services — fires BEFORE generic consulting so "services rendered" is caught specifically
    ("Services","Other Services","Needs Review",
     r"\b(services rendered|general service|misc(ellaneous)?\s+service|other service|labor service|service fee|service charge|professional fee)\b",
     "kw:svc_other"),
    # Generic consulting — fires last among Services patterns
    ("Services","Consulting & Advisory","Consulting",
     r"\b(consult(ing|ant)|advisory service|management consult|strategy consult|assessment service|evaluation service|technical assistance|professional service|contract service|program management service|project management service)\b",
     "kw:svc_consulting"),
    ("Printing, Marketing & Communications","Printing","Print",
     r"\b(print(ing)?|typeset|binding|poster|banner|brochure|flyer|letterhead)\b",
     "kw:printing"),
    ("Food & Catering","Food & Beverage","Catering",
     r"\b(cater(ing)?|food service|lunch|dinner|breakfast|refreshment|beverage|coffee service)\b",
     "kw:catering"),
    ("Utilities & Occupancy","Electric / Gas","Utilities",
     r"\b(electric(ity)?|utility bill|power bill|gas bill|water bill|kilowatt|kwh)\b",
     "kw:utilities"),
]

# ══════════════════════════════════════════════════════════════
# PASS 5 — ACCOUNT-FAMILY FALLBACK
# Source: VCU account_codes.csv — mapped by prefix families
# ══════════════════════════════════════════════════════════════
ACCOUNT_FAMILY_MAP = {
    # 600xxx — Contract services (broad)
    "600": ("Services","Consulting & Advisory","Account 600xxx Services"),
    # 620xxx — Supplies
    "620": ("Admin & Office","General Admin","Account 620xxx Supplies"),
    # 625xxx — Athletics
    "625": ("Travel, Events & Hospitality","Event / Venue","Account 625xxx Athletics"),
    # 630xxx — Telecom
    "630": ("IT","Network / Telecom","Account 630xxx Telecom"),
    # 634xxx — Commercial Consulting
    "634": ("Services","Consulting & Advisory","Account 634xxx Consulting"),
    # 635xxx — Postage
    "635": ("Admin & Office","Office Supplies","Account 635xxx Postage"),
    # 636xxx — Subgrants / Stipends
    "636": ("Inter-Entity / Transfers","Subgrants & Stipends","Account 636xxx Subgrant/Stipend"),
    # 637xxx — Leases / Software subscriptions
    "637": ("Utilities & Occupancy","Rent / Lease / Occupancy","Account 637xxx Leases"),
    # 638xxx — Travel, training, events, utilities, insurance, participant support
    "638": ("Travel, Events & Hospitality","Travel","Account 638xxx Travel/Events"),
    # 639xxx — Facilities Admin
    "639": ("Facilities / MRO","Trades Services","Account 639xxx Facilities Admin"),
    # 653xxx — Core Services (bioinformatics, cryo-EM, etc.)
    "653": ("Research & Laboratory","Scientific Services","Account 653xxx Core Services"),
    # 700xxx — Capitalized equipment
    "700": ("Research & Laboratory","Lab Equipment & Instruments","Account 700xxx Capital Equipment"),
    # 750xxx — Capital Project Indexes (construction, major A&E)
    "750": ("Capital Projects & Construction","General Contractor / CM","Account 750xxx Capital Project"),
    # 720xxx — Capital equipment / infrastructure
    "720": ("Capital Projects & Construction","Capital Equipment Installation","Account 720xxx Capital Equipment"),
    # 400xxx / 410xxx — Revenue / Fees (shouldn't appear in spend, but handle gracefully)
    "400": ("Admin & Office","Revenue / Fees","Account 400xxx Fees"),
    "410": ("Admin & Office","Revenue / Fees","Account 410xxx Fees"),
    "411": ("Admin & Office","General Admin","Account 411xxx Parking"),
}

_COMPILED = [
    (m, l2, l3, re.compile(p, re.IGNORECASE), hit)
    for m, l2, l3, p, hit in KEYWORD_PATTERNS
]


def categorize_row(row: dict) -> dict:
    vendor    = str(row.get("Vendor Name","")).strip().lower()
    commodity = str(row.get("Commodity Code","")).strip().lstrip("0") # normalize leading zeros
    commodity_raw = str(row.get("Commodity Code","")).strip()
    account   = str(row.get("Account","")).strip()
    cat1      = str(row.get("Category Level 1","")).strip().lower()
    desc      = str(row.get("Product Description","")).strip().lower()
    mfr       = str(row.get("Manufacturer","")).strip().lower()
    scan_text = f"{desc} {mfr} {cat1}"

    # Pass 0: vendor hard overrides (beats commodity code)
    for kw, (m, l2, l3, hit) in VENDOR_OVERRIDES.items():
        if kw in vendor:
            return _r(m, l2, l3, 0, hit, 0.95)

    # Pass 1: try both raw and leading-zero-stripped commodity
    for c in [commodity_raw, commodity]:
        if c in COMMODITY_CROSSWALK:
            m, l2, l3 = COMMODITY_CROSSWALK[c]
            return _r(m, l2, l3, 1, f"comm:{c}", 0.95)

    # Pass 2: vendor always-list
    for kw, (m, l2, l3, hit) in VENDOR_MAP.items():
        if kw in vendor:
            return _r(m, l2, l3, 2, hit, 0.9)

    # Pass 3: category level 1 metadata
    for kw, (m, l2, l3) in CATEGORY_L1_MAP.items():
        if kw in cat1:
            return _r(m, l2, l3, 3, f"cat1:{kw[:20]}", 0.7)

    # Pass 4: keyword / regex
    for m, l2, l3, pattern, hit in _COMPILED:
        if pattern.search(scan_text):
            return _r(m, l2, l3, 4, hit, 0.5)

    # Pass 5: account-family fallback
    for acct in account.replace("|", " ").split():
        pfx = acct[:3]
        if pfx in ACCOUNT_FAMILY_MAP:
            m, l2, l3 = ACCOUNT_FAMILY_MAP[pfx]
            return _r(m, l2, l3, 5, f"acct:{pfx}xxx", 0.2)

    return _r("Uncategorized", "Unknown", "Unknown", 5, "fallback:none", 0.1)


def _r(master, l2, l3, pass_num, hit, conf):
    return {"master_bucket": master, "sub_bucket_l2": l2, "sub_bucket_l3": l3,
            "rule_pass": pass_num, "rule_hit": hit, "confidence_score": conf}

# Services rows that need human review: landed here via account fallback only,
# OR matched only the generic consulting / other-services catch-all keywords.
_SERVICES_REVIEW_HITS = {"acct:600xxx", "acct:634xxx", "kw:svc_other", "kw:svc_consulting", "kw:svc_facilities"}

def _needs_services_review(row_result: dict) -> bool:
    """True if this Services row should be queued for manual sub-classification."""
    if row_result.get("master_bucket") != "Services":
        return False
    hit = row_result.get("rule_hit","")
    return (row_result.get("rule_pass",0) >= 5          # account-code fallback
            or hit in _SERVICES_REVIEW_HITS
            or hit.startswith("acct:"))


def confidence_label(score):
    if score >= 0.8: return "Very High"
    if score >= 0.6: return "High"
    if score >= 0.4: return "Medium"
    return "Low"

def rule_pass_label(n):
    return {0:"Vendor Hard Override", 1:"Commodity Code Crosswalk", 2:"Vendor Always-List",
            3:"Category Metadata", 4:"Keyword / Regex", 5:"Account-Family Fallback"}.get(n, "Unknown")

def categorize_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    # Run rule engine — one dict per row
    results = df.apply(lambda r: categorize_row(r.to_dict()), axis=1, result_type="expand")

    # Build all derived columns in one shot — avoids PerformanceWarning from incremental inserts
    extra = pd.DataFrame({
        "confidence_label":     results["confidence_score"].apply(confidence_label),
        "rule_pass_label":      results["rule_pass"].apply(rule_pass_label),
        "services_review_flag": results.apply(_needs_services_review, axis=1),
    }, index=results.index)

    return pd.concat([df.reset_index(drop=True),
                      results.reset_index(drop=True),
                      extra.reset_index(drop=True)], axis=1)
