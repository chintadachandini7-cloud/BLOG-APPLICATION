# Install python-docx if not already installed:
# pip install python-docx

import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

def create_report():
    doc = docx.Document()

    # Page Margins: 1 inch all around
    for section in doc.sections:
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)

    # Base Normal Style Settings
    style_normal = doc.styles['Normal']
    font = style_normal.font
    font.name = 'Times New Roman'
    font.size = Pt(12)
    font.color.rgb = RGBColor(0, 0, 0)
    style_normal.paragraph_format.line_spacing = 1.0
    style_normal.paragraph_format.space_after = Pt(6)

    def add_p(text="", align=WD_ALIGN_PARAGRAPH.LEFT, bold=False, italic=False, size=12, space_after=6, line_spacing=1.0):
        p = doc.add_paragraph()
        p.alignment = align
        p.paragraph_format.line_spacing = line_spacing
        p.paragraph_format.space_after = Pt(space_after)
        if text:
            run = p.add_run(text)
            run.font.name = 'Times New Roman'
            run.font.size = Pt(size)
            run.font.bold = bold
            run.font.italic = italic
        return p

    def add_h1(text):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(14)
        p.paragraph_format.space_after = Pt(6)
        p.paragraph_format.line_spacing = 1.0
        run = p.add_run(text.upper())
        run.font.name = 'Times New Roman'
        run.font.size = Pt(16)
        run.font.bold = True
        return p

    def add_h2(text):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(10)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.line_spacing = 1.0
        run = p.add_run(text)
        run.font.name = 'Times New Roman'
        run.font.size = Pt(14)
        run.font.bold = True
        return p

    def format_table(table):
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        for r_idx, row in enumerate(table.rows):
            for cell in row.cells:
                tcPr = cell._element.get_or_add_tcPr()
                # Light borders
                borders = parse_xml(r'<w:tcBorders %s><w:top w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"/><w:left w:val="none"/><w:right w:val="none"/></w:tcBorders>' % nsdecls('w'))
                tcPr.append(borders)
                for p in cell.paragraphs:
                    p.paragraph_format.line_spacing = 1.15
                    p.paragraph_format.space_after = Pt(2)
                    for r in p.runs:
                        r.font.name = 'Times New Roman'
                        r.font.size = Pt(12)
                        if r_idx == 0:
                            r.font.bold = True
                            p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # ================= COVER PAGE =================
    add_p("SMART PUBLIC TRANSPORT TRACKING & MANAGEMENT SYSTEM", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=18, space_after=12)
    add_p("Community Project Report", align=WD_ALIGN_PARAGRAPH.CENTER, italic=True, size=14, space_after=18)
    add_p("Submitted by", align=WD_ALIGN_PARAGRAPH.CENTER, size=12, space_after=10)

    # 2x2 Student Box Table
    t_students = doc.add_table(rows=2, cols=2)
    t_students.alignment = WD_TABLE_ALIGNMENT.CENTER
    members = [
        ("CHINTADA CHANDINI", "24331A4414"),
        ("GUDALA JASIN", "24331A4422"),
        ("SAMANTHULA MANEENDRA", "24331A4454"),
        ("YEDLA KARTHIK", "24331A4465")
    ]
    for idx, (name, reg) in enumerate(members):
        r, c = idx // 2, idx % 2
        p = t_students.cell(r, c).paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run1 = p.add_run(f"{name}\n")
        run1.font.bold = True
        run1.font.size = Pt(12)
        run2 = p.add_run(f"({reg})")
        run2.font.size = Pt(11)
    
    add_p("", space_after=14)
    add_p("In partial fulfillment for the award of the degree\nof", align=WD_ALIGN_PARAGRAPH.CENTER, size=12, space_after=4)
    add_p("BACHELOR OF TECHNOLOGY\nIN\nCOMPUTER SCIENCE & ENGINEERING\n(Artificial Intelligence & Machine Learning)", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=13, space_after=14)
    add_p("Under the esteemed Guidance of", align=WD_ALIGN_PARAGRAPH.CENTER, size=12, space_after=2)
    add_p("Dr. H. Y. PRASANNA RAJU\nAssociate Professor", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=12, space_after=16)
    
    add_p("DEPARTMENT OF DATA ENGINEERING\nMAHARAJ VIJAYARAM GAJAPATHI RAJ COLLEGE OF ENGINEERING (Autonomous)\n(Approved by AICTE, New Delhi, and permanently affiliated to JNTUGV, Vizianagaram),\nListed u/s 2(f) & 12(B) of UGC Act 1956.\nVijayaram Nagar Campus, Chintalavalasa, Vizianagaram-535005, Andhra Pradesh\nOctober, 2025", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=10, space_after=0)

    doc.add_page_break()

    # ================= CERTIFICATE =================
    add_h1("CERTIFICATE")
    p_cert = add_p(
        "This is to certify that the project entitled “SMART PUBLIC TRANSPORT” is the bonafide work carried out by "
        "Chintada Chandini (24331A4414), Gudala Jasin (24331A4422), Samanthula Maneendra (24331A4454), and Yedla Karthik (24331A4465), "
        "of B.Tech V Sem CSE-AIML, M.V.G.R. College of Engineering (Autonomous), Vizianagaram, during the year 2025-2026, "
        "in partial fulfilment of the requirements for the award of the Degree of Bachelor of Technology and that the project "
        "has not formed the basis for the award previously of any degree or any other similar title.",
        align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_after=40
    )

    t_sigs = doc.add_table(rows=1, cols=2)
    t_sigs.alignment = WD_TABLE_ALIGNMENT.CENTER
    p_left = t_sigs.cell(0, 0).paragraphs[0]
    p_left.add_run("Signature of Project Guide\n\n\n\nDr. H. Y. Prasanna Raju\nAssociate Professor\nDepartment: Data Engineering").font.bold = True
    p_right = t_sigs.cell(0, 1).paragraphs[0]
    p_right.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_right.add_run("Signature of Head of the Department\n\n\n\nDr. V. Jyothi\nAssociate Professor\nDepartment: Data Engineering").font.bold = True

    doc.add_page_break()

    # ================= DECLARATION =================
    add_h1("DECLARATION")
    add_p(
        "We hereby declare that the work done on the dissertation entitled “SMART PUBLIC TRANSPORT” has been carried out by us "
        "and submitted in partial fulfilment for the award of credits in Bachelor of Technology in Computer Science and Engineering "
        "(Artificial Intelligence & Machine Learning) of M.V.G.R College of Engineering (Autonomous) and affiliated to JNTUGV, Vizianagaram. "
        "The various contents incorporated in the dissertation have not been submitted for the award of any degree of any other institution or university.",
        align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_after=30
    )
    for name, reg in members:
        add_p(f"{name}  ({reg})", align=WD_ALIGN_PARAGRAPH.RIGHT, bold=True, space_after=4)

    doc.add_page_break()

    # ================= ACKNOWLEDGEMENT =================
    add_h1("ACKNOWLEDGEMENT")
    add_p(
        "We express our sincere gratitude to Dr. H. Y. Prasanna Raju for his invaluable guidance and support as our mentor throughout the project. "
        "His unwavering commitment to excellence and constructive feedback motivated us to achieve our project goals. We are greatly indebted to him for his exceptional guidance.",
        align=WD_ALIGN_PARAGRAPH.JUSTIFY
    )
    add_p(
        "Additionally, we extend our thanks to Prof. P. S. Sitharama Raju (Director), Dr. Y. M. C. Shekar (Principal), and Dr. V. Jyothi (Head of the Department) "
        "for their unwavering support and assistance, which were instrumental in the successful completion of the project. We are thankful for and fortunate enough "
        "to get constant encouragement and guidance from our Project Coordinator, Mrs. G. Gayathri.",
        align=WD_ALIGN_PARAGRAPH.JUSTIFY
    )
    add_p(
        "We also acknowledge the dedicated assistance provided by all the staff members in the Department of Data Engineering. "
        "Finally, we appreciate the contributions of all those who directly or indirectly contributed to the successful execution of this endeavor.",
        align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_after=24
    )
    for name, reg in members:
        add_p(f"{name} ({reg})", align=WD_ALIGN_PARAGRAPH.RIGHT, bold=True, space_after=3)

    doc.add_page_break()

    # ================= ABSTRACT =================
    add_h1("ABSTRACT")
    add_p(
        "Public transportation in small cities often lacks real-time tracking, making it difficult for passengers to know bus locations and arrival times. "
        "Public transit systems in small cities and towns rely heavily on static, fixed schedules that fail to reflect real-world traffic, delays, or route diversions. "
        "Due to the lack of real-time location tracking and estimated time of arrival (ETA) systems, commuters face unpredictable waiting times, reduced travel efficiency, "
        "and overall inconvenience, ultimately diminishing the reliability of local public transportation. Existing bus tracking systems in major metro cities rely on high-cost "
        "Intelligent Transportation Systems (ITS) and heavy hardware infrastructure that small municipalities cannot afford.",
        align=WD_ALIGN_PARAGRAPH.JUSTIFY
    )
    add_p(
        "The project develops a web-based application providing live bus tracking, estimated arrival time (ETA), route details, and service updates through API integration. "
        "The system interconnects four distinct user roles: Passengers, Drivers, Conductors, and Transport Administrators. The backend architecture is built with Python Flask "
        "and exposes a REST API layer interfacing with SQLite (transport.db) and Supabase Realtime for cloud synchronization. By capturing driver/conductor device coordinates and mapping "
        "routes with Leaflet.js and OpenStreetMap, the system displays dynamic bus markers, stop sequences, crowd indicators, and Emergency SOS feeds to improve transit reliability without dedicated on-board GPS hardware.",
        align=WD_ALIGN_PARAGRAPH.JUSTIFY
    )

    doc.add_page_break()

    # ================= TABLE OF CONTENTS =================
    add_h1("TABLE OF CONTENTS")
    t_toc = doc.add_table(rows=1, cols=2)
    t_toc.rows[0].cells[0].paragraphs[0].add_run("Contents")
    t_toc.rows[0].cells[1].paragraphs[0].add_run("Page No.")
    
    toc_items = [
        ("List of Abbreviations", "i"),
        ("List of Figures", "ii"),
        ("List of Tables", "iii"),
        ("1. Introduction", "1"),
        ("    1.1 Problem Statement", "1"),
        ("    1.2 Project Objective", "2"),
        ("    1.3 Scope of the Project", "2"),
        ("2. Literature Survey", "3"),
        ("3. Data Gathering / Data Used", "4"),
        ("4. Methodology / System Design", "5"),
        ("5. Implementation / Modules", "7"),
        ("6. Results / Outputs", "9"),
        ("7. Impact Assessment", "11"),
        ("8. Challenges Faced", "12"),
        ("9. Conclusion", "13"),
        ("10. Future Work", "14"),
        ("References", "15"),
        ("Appendix A: Packages, Tools Used & Working Process", "16"),
        ("Appendix B: Source Code", "17"),
        ("Paper Publications (if any)", "19")
    ]
    for text, page in toc_items:
        row = t_toc.add_row()
        row.cells[0].paragraphs[0].add_run(text)
        p_p = row.cells[1].paragraphs[0]
        p_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p_p.add_run(page)
    format_table(t_toc)

    doc.add_page_break()

    # ================= LIST OF ABBREVIATIONS, FIGURES & TABLES =================
    add_h1("LIST OF ABBREVIATIONS")
    abbreviations = [
        ("API", "Application Programming Interface"),
        ("APSRTC", "Andhra Pradesh State Road Transport Corporation"),
        ("ATA", "Actual Time of Arrival"),
        ("CRUD", "Create, Read, Update, Delete"),
        ("ETA", "Estimated Time of Arrival"),
        ("GPS", "Global Positioning System"),
        ("GTFS", "General Transit Feed Specification"),
        ("ITS", "Intelligent Transportation Systems"),
        ("OSM", "OpenStreetMap"),
        ("RBAC", "Role-Based Access Control"),
        ("REST", "Representational State Transfer"),
        ("SOS", "Save Our Souls (Emergency Distress Trigger)"),
        ("TGSRTC", "Telangana State Road Transport Corporation"),
        ("UI", "User Interface")
    ]
    for abb, full in abbreviations:
        p = add_p()
        p.add_run(f"{abb}: ").bold = True
        p.add_run(full)

    add_h1("LIST OF FIGURES")
    t_lof = doc.add_table(rows=1, cols=2)
    t_lof.rows[0].cells[0].paragraphs[0].add_run("Figure No. & Title")
    t_lof.rows[0].cells[1].paragraphs[0].add_run("Page No.")
    figures = [
        ("Figure 4.1: System Execution Flow Chart (4 User Roles)", "5"),
        ("Figure 4.2: Block Diagram of Multi-tier System Architecture", "6"),
        ("Figure 6.1: Passenger Route Search & Discovery Interface", "9"),
        ("Figure 6.2: Route Stop Sequences and Dynamic Fare Display", "10"),
        ("Figure 6.3: Emergency SOS Dispatcher & Platform Safety Feed Screen", "10")
    ]
    for fig, p_no in figures:
        row = t_lof.add_row()
        row.cells[0].paragraphs[0].add_run(fig)
        p = row.cells[1].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p.add_run(p_no)
    format_table(t_lof)

    add_h1("LIST OF TABLES")
    t_lot = doc.add_table(rows=1, cols=2)
    t_lot.rows[0].cells[0].paragraphs[0].add_run("Table No. & Title")
    t_lot.rows[0].cells[1].paragraphs[0].add_run("Page No.")
    tables_list = [
        ("Table 2.1: Comparative Analysis of Existing Transit Systems vs Proposed Solution", "3"),
        ("Table 3.1: Data Schema Overview for Static and Real-Time Transit Entities", "4"),
        ("Table 5.1: Core REST Endpoints Defined in Flask Engine", "8")
    ]
    for tab, p_no in tables_list:
        row = t_lot.add_row()
        row.cells[0].paragraphs[0].add_run(tab)
        p = row.cells[1].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p.add_run(p_no)
    format_table(t_lot)

    doc.add_page_break()

    # ================= 1. INTRODUCTION =================
    add_h1("1. INTRODUCTION")
    add_p(
        "The Smart Public Transport Tracking for Small Cities project is a web-based application developed to modernize public transport services in small cities and towns. "
        "It provides passengers with an easy-to-use platform to access transport-related information through a digital interface. "
        "By integrating web technologies and APIs, the system improves communication between transport services and passengers, making public transportation more efficient, reliable, and user-friendly.",
        align=WD_ALIGN_PARAGRAPH.JUSTIFY
    )
    add_h2("1.1 Problem Statement")
    add_p(
        "Public transit systems in small cities and towns rely heavily on static, fixed schedules that fail to reflect real-world traffic, delays, or route diversions. "
        "Due to the lack of real-time location tracking and estimated time of arrival (ETA) systems, commuters face unpredictable waiting times, reduced travel efficiency, "
        "and overall inconvenience, ultimately diminishing the reliability of local public transportation.",
        align=WD_ALIGN_PARAGRAPH.JUSTIFY
    )
    add_h2("1.2 Project Objective")
    objectives = [
        "Develop a web-based bus tracking system.",
        "Provide real-time bus location and Estimated Time of Arrival (ETA).",
        "Display accurate route details and bus stop information.",
        "Provide a simple, user-friendly, and responsive interface.",
        "Implement a unified management ecosystem serving Passengers, Drivers, Conductors, and Transit Administrators."
    ]
    for obj in objectives:
        p = add_p(f"•  {obj}")
        p.paragraph_format.left_indent = Inches(0.25)
    
    add_h2("1.3 Scope of the Project")
    add_p(
        "The project is limited to developing a web-based application for public transport tracking in small cities and towns. "
        "It provides bus location, estimated arrival time (ETA), route information, and service notifications through API integration. "
        "The system does not include GPS hardware installation, ticket booking, online payment, or direct control of bus operations. "
        "Its primary focus is to improve passenger access to transport information and support better travel planning.",
        align=WD_ALIGN_PARAGRAPH.JUSTIFY
    )

    # ================= 2. LITERATURE SURVEY =================
    add_h1("2. LITERATURE SURVEY")
    add_p(
        "Existing bus tracking systems in major metro cities rely on high-cost Intelligent Transportation Systems (ITS) and heavy hardware infrastructure. "
        "Smaller cities and towns cannot afford these expensive setups and still depend on outdated, fixed paper schedules. "
        "These rigid schedules fail to reflect real-world traffic jams, vehicle breakdowns, or route changes. "
        "Recent advances in web technologies, mobile GPS, and open APIs make real-time data streaming much cheaper and easier to implement. "
        "This project fills the existing gap by offering a low-cost, lightweight web application tailored specifically for small-town transit networks.",
        align=WD_ALIGN_PARAGRAPH.JUSTIFY
    )
    
    add_p("Table 2.1: Comparative Analysis of Existing Transit Systems vs Proposed Solution", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True)
    t_lit = doc.add_table(rows=1, cols=4)
    t_lit.rows[0].cells[0].paragraphs[0].add_run("Parameters")
    t_lit.rows[0].cells[1].paragraphs[0].add_run("Metro ITS Systems")
    t_lit.rows[0].cells[2].paragraphs[0].add_run("Static RTC Portals")
    t_lit.rows[0].cells[3].paragraphs[0].add_run("Proposed System")
    lit_data = [
        ("Hardware Setup", "Custom On-Board Units / Telemetry", "Central Web Servers only", "Smartphone Browser Geolocation"),
        ("Live Tracking", "Available via dedicated GPS pings", "Not Available (Timetable only)", "Available via WebSockets & OpenStreetMap"),
        ("Deployment Cost", "High Capital & Maintenance", "Low (Static server costs)", "Very Low (No hardware retrofit needed)"),
        ("Passenger Safety", "Limited to manual SMS", "Not Integrated", "Built-in Emergency SOS & Corridor Safety Feed")
    ]
    for p1, p2, p3, p4 in lit_data:
        r = t_lit.add_row()
        r.cells[0].paragraphs[0].add_run(p1)
        r.cells[1].paragraphs[0].add_run(p2)
        r.cells[2].paragraphs[0].add_run(p3)
        r.cells[3].paragraphs[0].add_run(p4)
    format_table(t_lit)

    # ================= 3. DATA GATHERING =================
    add_h1("3. DATA GATHERING / DATA USED")
    add_p(
        "The project integrates static public transport timetables with dynamic real-time telemetry captured through web APIs. "
        "Official timetable datasets from APSRTC and TGSRTC were ingested to establish the baseline route topologies, stop milestones, and service classifications.",
        align=WD_ALIGN_PARAGRAPH.JUSTIFY
    )
    add_p("Table 3.1: Data Schema Overview for Static and Real-Time Transit Entities", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True)
    t_data = doc.add_table(rows=1, cols=3)
    t_data.rows[0].cells[0].paragraphs[0].add_run("Data Category")
    t_data.rows[0].cells[1].paragraphs[0].add_run("Entity Source")
    t_data.rows[0].cells[2].paragraphs[0].add_run("Attributes & Field Specifications")
    data_items = [
        ("Fleet Profile", "Depot Records / Admin CRUD", "bus_number, bus_name, route_id, status, delay_status, delay_minutes"),
        ("Route Topologies", "APSRTC / TGSRTC Timetables", "route_name, source, destination, operator, service_type, data_source"),
        ("Stop Milestones", "OpenStreetMap / Transit Data", "route_id, stop_name, latitude, longitude, stop_order, scheduled_arrival"),
        ("Arrival Telemetry", "Runtime Engine Calculation", "bus_id, stop_id, ETA (Estimated), ATA (Actual), recorded_at"),
        ("Live Geolocation", "Browser Geolocation API", "trip_id, bus_id, latitude, longitude, accuracy, timestamp")
    ]
    for c1, c2, c3 in data_items:
        r = t_data.add_row()
        r.cells[0].paragraphs[0].add_run(c1)
        r.cells[1].paragraphs[0].add_run(c2)
        r.cells[2].paragraphs[0].add_run(c3)
    format_table(t_data)

    # ================= 4. METHODOLOGY =================
    add_h1("4. METHODOLOGY / SYSTEM DESIGN")
    add_p(
        "The architecture is organized into four distinct role tiers (Passenger, Driver, Conductor, Administrator) communicating "
        "with a central Python Flask backend through RESTful APIs and synchronized with SQLite and Supabase Realtime layers.",
        align=WD_ALIGN_PARAGRAPH.JUSTIFY
    )
    add_h2("4.1 Role-Based Access Control and Execution Flow")
    roles = [
        ("Passenger Workflow: ", "Accesses the passenger dashboard without mandatory credentials. Searches for origins and destinations, inspects intermediate stops, and watches live vehicle markers, ETAs, and delay alerts on a Leaflet map."),
        ("Driver Workflow: ", "Logs in with phone number and password. Upon session verification, selects an assigned bus, starts the trip, and pushes device geolocation coordinates continuously."),
        ("Conductor Workflow: ", "Authenticates, joins active vehicle trips, pushes stop clearance delay alerts, updates crowd density levels, and terminates the trip session."),
        ("Administrator Workflow: ", "Guarded by strict session checks. Manages buses, routes, stops, and transit staff credentials.")
    ]
    for title, desc in roles:
        p = add_p()
        p.add_run(f"•  {title}").bold = True
        p.add_run(desc)
        p.paragraph_format.left_indent = Inches(0.25)

    # ================= 5. IMPLEMENTATION =================
    add_h1("5. IMPLEMENTATION / MODULES")
    add_p(
        "The application logic is driven by app.py (containing ~1,800 lines of Python code) interfacing with database.py to maintain "
        "relational integrity across buses, routes, stops, trips, and live_locations.",
        align=WD_ALIGN_PARAGRAPH.JUSTIFY
    )
    add_p("Table 5.1: Core REST Endpoints Defined in Flask Engine", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True)
    t_api = doc.add_table(rows=1, cols=4)
    t_api.rows[0].cells[0].paragraphs[0].add_run("HTTP Method")
    t_api.rows[0].cells[1].paragraphs[0].add_run("Endpoint")
    t_api.rows[0].cells[2].paragraphs[0].add_run("Authorized Role")
    t_api.rows[0].cells[3].paragraphs[0].add_run("Operational Responsibility")
    api_items = [
        ("GET", "/api/health", "Public", "Performs database connectivity and server status checks"),
        ("GET, POST", "/api/routes", "Public / Admin", "Retrieves active transit corridors or registers new routes"),
        ("GET, POST", "/api/buses", "Public / Admin", "Queries fleet records or adds new bus profiles"),
        ("GET", "/api/bus/by-number/<num>", "Public", "Fetches current status, delay minutes, and GPS coords"),
        ("POST", "/api/driver/start-trip", "Driver", "Validates session, sets trip status to ACTIVE, assigns bus"),
        ("POST", "/api/conductor/join-trip", "Conductor", "Attaches authenticated conductor to an active trip"),
        ("POST", "/api/conductor/leave-trip", "Conductor", "Detaches conductor from the trip record")
    ]
    for c1, c2, c3, c4 in api_items:
        r = t_api.add_row()
        r.cells[0].paragraphs[0].add_run(c1)
        r.cells[1].paragraphs[0].add_run(c2)
        r.cells[2].paragraphs[0].add_run(c3)
        r.cells[3].paragraphs[0].add_run(c4)
    format_table(t_api)

    # ================= 6. RESULTS =================
    add_h1("6. RESULTS / OUTPUTS")
    add_p(
        "The system has been evaluated along the high-traffic corridor from Vijayawada (Pandit Nehru Bus Station) to Visakhapatnam (Dwaraka Bus Station):",
        align=WD_ALIGN_PARAGRAPH.JUSTIFY
    )
    res_points = [
        "Multilingual Commuter UI: Seamless toggle between English, Telugu, and Hindi interfaces with route search and discovery.",
        "Interactive Journey Tracker: Real-time stop progression visualization across Guntur, Kakinada, and Nellore junctions with dynamic ETAs.",
        "Crowd Density Indicators: Visual cues showing vehicle passenger occupancy (Low, Medium, High).",
        "Corridor Safety Module: A dedicated safety view featuring an Emergency SOS broadcast button and local safety alerts."
    ]
    for pt in res_points:
        p = add_p(f"•  {pt}")
        p.paragraph_format.left_indent = Inches(0.25)

    # ================= 7. IMPACT ASSESSMENT =================
    add_h1("7. IMPACT ASSESSMENT")
    add_h2("Technical Impact")
    add_p("Provides real-time transport information and efficient API integration without requiring capital-intensive proprietary on-board telematics hardware.")
    add_h2("Social Impact")
    add_p("Reduces passenger waiting time, eliminates scheduling ambiguity at rural and semi-urban stops, and improves women and night traveler safety via distress feeds.")
    add_h2("Economic Impact")
    add_p("Encourages public transport usage, saves passenger travel time, and reduces operational overhead for regional depots.")

    # ================= 8. CHALLENGES FACED =================
    add_h1("8. CHALLENGES FACED")
    challenges = [
        "Availability of real-time transport data in regional sectors.",
        "API integration across disparate data formats and coordinate systems.",
        "GPS accuracy degradation inside metallic vehicle cabins.",
        "Data synchronization between local SQLite stores and cloud Supabase endpoints.",
        "Testing and validating telemetry tracking on moving transit routes."
    ]
    for ch in challenges:
        p = add_p(f"•  {ch}")
        p.paragraph_format.left_indent = Inches(0.25)

    # ================= 9. CONCLUSION =================
    add_h1("9. CONCLUSION")
    add_p(
        "The Smart Public Transport Tracking for Small Cities project provides an efficient web-based solution for improving public transportation through "
        "real-time bus tracking, estimated arrival times (ETA), and route information. By integrating transport data through APIs, the system helps passengers "
        "plan their journeys more effectively, reduces waiting time, and enhances the overall travel experience. The project demonstrates how digital technologies "
        "can make public transportation more reliable, accessible, and efficient for both passengers and transport authorities.",
        align=WD_ALIGN_PARAGRAPH.JUSTIFY
    )

    # ================= 10. FUTURE WORK =================
    add_h1("10. FUTURE WORK")
    future = [
        "Expansion to more cities and transport services.",
        "Online ticket booking and digital payments.",
        "Push notifications and automated service alerts.",
        "Analytics dashboard for transport authorities.",
        "AI-based arrival prediction using historical delay factors.",
        "Implementation in keypad mobiles through SMS and USSD gateways."
    ]
    for fw in future:
        p = add_p(f"•  {fw}")
        p.paragraph_format.left_indent = Inches(0.25)

    # ================= REFERENCES =================
    add_h1("REFERENCES")
    refs = [
        "[1] APSRTC Online Passenger Reservation System Timetable Data: https://www.apsrtconline.in/oprs-web/services/timeTable.do",
        "[2] TGSRTC Official Timetable Services: https://www.tgsrtc.telangana.gov.in/",
        "[3] OpenStreetMap & Leaflet.js Mapping Services: https://www.openstreetmap.org/#map=4/21.84/82.79",
        "[4] Development Repository & Environment: https://solid-enigma-x57x4w74g77w2pg9v.github.dev/"
    ]
    for r in refs:
        add_p(r, space_after=3)

    # ================= APPENDIX A =================
    add_h1("APPENDIX A: PACKAGES, TOOLS USED & WORKING PROCESS")
    add_h2("Packages & Tools Used")
    tools = [
        "Backend Framework: Python 3.10+, Flask",
        "Database Engines: SQLite3 (transport.db), Supabase Realtime (PostgreSQL)",
        "Frontend: HTML5, CSS3, JavaScript (ES6+), React, Tailwind CSS",
        "Mapping Engine: Leaflet.js, OpenStreetMap API",
        "Environment: Visual Studio Code, Git, GitHub, Vercel"
    ]
    for t in tools:
        p = add_p(f"•  {t}")
        p.paragraph_format.left_indent = Inches(0.25)
    
    add_h2("Working Process")
    add_p(
        "1. Ingest official APSRTC/TGSRTC bus routes and stop coordinates using import_official_data.py.\n"
        "2. Initialize relational database tables (buses, routes, stops, trips, live_locations) via database.py.\n"
        "3. Implement RESTful API endpoints and session-based role verification in app.py.\n"
        "4. Construct mobile-responsive views for Passengers, Drivers, Conductors, and Administrators.\n"
        "5. Connect Leaflet.js map layers with live coordinate feeds and emergency safety feeds."
    )

    # ================= APPENDIX B =================
    add_h1("APPENDIX B: SOURCE CODE")
    add_h2("Database Initialization Schema (database.py)")
    code_db = (
        "import sqlite3\n\n"
        "def init_db():\n"
        "    conn = sqlite3.connect('transport.db')\n"
        "    c = conn.cursor()\n"
        "    c.execute('''CREATE TABLE IF NOT EXISTS buses (\n"
        "        id INTEGER PRIMARY KEY AUTOINCREMENT,\n"
        "        bus_number TEXT UNIQUE NOT NULL,\n"
        "        bus_name TEXT, route_id INTEGER,\n"
        "        current_latitude REAL, current_longitude REAL,\n"
        "        status TEXT DEFAULT 'INACTIVE',\n"
        "        delay_status TEXT DEFAULT 'ON_TIME',\n"
        "        delay_minutes INTEGER DEFAULT 0\n"
        "    )''')\n"
        "    c.execute('''CREATE TABLE IF NOT EXISTS trips (\n"
        "        id INTEGER PRIMARY KEY AUTOINCREMENT,\n"
        "        bus_id INTEGER, driver_id INTEGER, conductor_id INTEGER,\n"
        "        route_id INTEGER, start_time TEXT, end_time TEXT,\n"
        "        status TEXT DEFAULT 'SCHEDULED'\n"
        "    )''')\n"
        "    c.execute('''CREATE TABLE IF NOT EXISTS live_locations (\n"
        "        id INTEGER PRIMARY KEY AUTOINCREMENT,\n"
        "        trip_id INTEGER, bus_id INTEGER,\n"
        "        latitude REAL, longitude REAL, accuracy REAL, timestamp TEXT\n"
        "    )''')\n"
        "    conn.commit()\n"
        "    conn.close()\n"
    )
    p_code1 = add_p(code_db, size=9.5)
    p_code1.runs[0].font.name = 'Consolas'

    add_h2("Trip Initiation Endpoint (app.py)")
    code_app = (
        "@app.route('/api/driver/start-trip', methods=['POST'])\n"
        "def start_trip():\n"
        "    if session.get('user_role') != 'driver':\n"
        "        return jsonify({'error': 'Unauthorized'}), 403\n"
        "    data = request.get_json()\n"
        "    driver_id = session.get('driver_id')\n"
        "    bus_id = data.get('bus_id')\n"
        "    route_id = data.get('route_id')\n"
        "    conn = get_db_connection()\n"
        "    cursor = conn.cursor()\n"
        "    cursor.execute('''INSERT INTO trips (bus_id, driver_id, route_id, start_time, status)\n"
        "                      VALUES (?, ?, ?, datetime('now'), 'ACTIVE')''', (bus_id, driver_id, route_id))\n"
        "    trip_id = cursor.lastrowid\n"
        "    cursor.execute('UPDATE buses SET status = \"ACTIVE\" WHERE id = ?', (bus_id,))\n"
        "    conn.commit()\n"
        "    conn.close()\n"
        "    return jsonify({'status': 'ACTIVE', 'trip_id': trip_id})\n"
    )
    p_code2 = add_p(code_app, size=9.5)
    p_code2.runs[0].font.name = 'Consolas'

    # ================= PAPER PUBLICATIONS =================
    add_h1("PAPER PUBLICATIONS (IF ANY)")
    add_p("Status: Under preparation for submission to student technical symposiums / conferences.", italic=True)

    filename = "Community_Project_Report_Smart_Public_Transport.docx"
    doc.save(filename)
    print(f"Report generated: {filename}")

if __name__ == '__main__':
    create_report()