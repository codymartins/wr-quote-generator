'''
Quote Generator for Waste Robotics
Author: Cody Martins
'''
import streamlit as st
from PIL import Image
import pandas as pd
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
import os

# --- WR Branding Setup ---
logo = Image.open("logoWasteRobotics(1).png")  # Make sure this file is in the same directory
col1, col2 = st.columns([1, 6])
with col1:
    st.image("logoWasteRobotics(1).png", width=80)
with col2:
    st.markdown("<h1 style='color: white;'>Waste Robotics Quote Generator</h1>", unsafe_allow_html=True)
    st.markdown("<p style='color: #EF3A2D; font-style: italic;'>Smarter Sorting with Robotics</p>", unsafe_allow_html=True)

st.markdown("""
    <style>
        .reportview-container {
            background-color: #0F0F0F;
            color: white;
        }
        .sidebar .sidebar-content {
            background-color: #1A1A1A;
        }
        h1, h2, h3 {
            color: #EF3A2D;
        }
        .stButton>button {
            background-color: #EF3A2D;
            color: white;
            border: none;
            padding: 0.5em 1em;
            font-weight: bold;
        }
        .stButton>button:hover {
            background-color: #c23024;
        }
        footer {
            visibility: hidden;
        }
        .footer {
            position: fixed;
            bottom: 0;
            width: 100%;
            background-color: #1A1A1A;
            color: white;
            text-align: center;
            padding: 5px;
            font-size: 12px;
        }
    </style>
    """, unsafe_allow_html=True)

# Load pricing data
pricing_df = pd.read_csv("pricing.csv")

# Convert it to a dictionary for easy lookup
PRICING = dict(zip(pricing_df["item"], pricing_df["price_usd"]))

robot_price = PRICING["robot_arm"]
gripper_price = PRICING["gripper"]
conveyor_price = PRICING["conveyor"]
hypervision_scanner_price = PRICING["hypervision_scanner"]
ai_training_price = PRICING["ai_training"]
software_license_price = PRICING["software_license"]
fat_price = PRICING["fat"]
try_and_buy_arm_price = PRICING["try_and_buy_arm"]
installation_price = PRICING["installation"]
shipping_per_km_price = PRICING["shipping_per_km"]
modification_small_price = PRICING['modification_small']
modification_medium_price = PRICING['modification_medium']
modification_large_price = PRICING['modification_large']
security_perimeter_price = PRICING['security_perimeter']
warranty_price = PRICING['warranty']
pe_stamp_price = PRICING['pe_stamp']
sat_price = PRICING['sat']

# --- Constants ---
CURRENCY_CONVERSION = {"USD": 1.0, "CAD": 1.36, "EUR": 0.92}
# PRICING = {
#     "robot_arm": 45000,
#     "gripper": 6000,
#     "conveyor": 8000,
#     "hypervision_scanner": 12000,
#     "ai_training": 15000,
#     "software_license": 5000,
#     "fat": 4000,
#     "try_and_buy_arm": 90000,
#     "installation": 30000,
#     "shipping_per_km": 3.5,
#     "modification_small": 5000,
#     "modification_medium": 15000,
#     "modification_large": 30000,
#     "security_perimeter": 2500,
#     "warranty": 8000,
#     "pe_stamp": 2000,
#     "sat": 6000
# }

# --- UI ---
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "Proposal Info", 
    "System Config", 
    "Technical Specs", 
    "Shipping & Timeline", 
    "Inclusions & Quote"
])

with tab1:
    st.header("Proposal Information")
    st.progress(20, text="Step 1 of 5")
    quote_date = st.date_input("Quote Date")
    value_proposition = st.text_input("Value Proposition (Main Proposal Title)")
    client_name = st.text_input("Client Name")
    client_company = st.text_input("Client Company Name")
    salesman_name = st.text_input("Salesperson Name")
    site_location = st.text_input("Site Location")
    currency = st.selectbox("Currency", ["USD", "CAD", "EUR"])
    application_overview = st.text_area("Brief Summary of the Application")

with tab2:
    st.header("System Configuration")
    st.progress(40, text="Step 2 of 5")
    materials = st.multiselect("Materials to Sort", ["PCBs", "UBCs", "Trash", "Other"])
    try_and_buy = st.checkbox("Include Try & Buy Option?")
    conveyor_included = st.selectbox("Conveyor Provided?", ["Yes", "No"])
    conveyor_size = st.text_input("Conveyor Size (inches)", disabled=(conveyor_included == "No"))
    belt_speed = st.text_input("Belt Speed (e.g. 80 ft/min)")
    pick_rate = st.text_input("Pick Rate (e.g. 35 picks/minute)")
    modification_scale = st.selectbox("Modification Scope", ["Small", "Medium", "Large"])
    hypervision_scanners = st.number_input("Number of Hypervision Scanners", min_value=0, value=1)
    robot_type = st.selectbox("Robot Type", ["Fanuc LR-Mate 200ID", "Other"])
    robot_arms = st.slider("Number of Robot Arms", 1, 3, 2)
    gripper_type = st.multiselect("Gripper Type", ["VentuR (suction)", "PinchR", "Other"])

with tab3:
    st.header("Technical Specs")
    st.progress(60, text="Step 3 of 5")
    max_object_weight = st.number_input("Maximum Object Weight per Robot (kg)", min_value=0.0)
    robot_bases = st.number_input("Number of Robot Bases", min_value=0)
    vision_system = st.multiselect("Robot Vision System", ["RGB Camera", "3D Camera"])
    input_power_kva = st.number_input("Input Power (kVA)", min_value=0.0)
    avg_consumption_kw = st.number_input("Average Power Consumption (kW)", min_value=0.0)
    air_consumption_lpm = st.number_input("Total Air Consumption (L/min)", min_value=0)

with tab4:
    st.header("Shipping & Timeline")
    st.progress(80, text="Step 4 of 5")
    shipping_distance = st.text_input("Estimated Shipping Distance (miles or km)")
    timeline = {
        "Order Confirmation / Project Kickoff": st.text_input("Order Confirmation / Project Kickoff Duration"),
        "Detailed Engineering": st.text_input("Detailed Engineering Duration"),
        "Engineering Review": st.text_input("Engineering Review Duration"),
        "Procurement and Fabrication": st.text_input("Procurement and Fabrication Duration"),
        "FAT and Shipping": st.text_input("FAT and Shipping Duration"),
        "Retrofit and Installation": st.text_input("Retrofit and Installation Duration"),
        "Commissioning and SAT": st.text_input("Commissioning and SAT Duration")
    }

with tab5:
    st.header("Inclusions & Final Quote")
    st.progress(100, text="Step 5 of 5")
    security_perimeter = st.checkbox("Include Security Perimeter?")
    software_license = st.checkbox("Include Perpetual Software License?")
    ai_training = st.checkbox("Include AI Custom Training?")
    warranty = st.checkbox("Include Warranty?")
    fat = st.checkbox("Include Factory Pre-Assembly and Testing (FAT)?")
    pe_stamp = st.checkbox("Include PE Stamp?")
    sat = st.checkbox("Include Site Acceptance Test (SAT)?")
    installation_by_wr = st.selectbox("Will WR handle site modifications and installation?", ["Yes", "No", "Partially"])

    if st.button("Generate Quote"): 
        # --- Input Validation ---
        missing_fields = []

        # Proposal Info
        if not value_proposition.strip():
            missing_fields.append("Value Proposition")
        if not client_name.strip():
            missing_fields.append("Client Name")
        if not client_company.strip():
            missing_fields.append("Client Company Name")
        if not salesman_name.strip():
            missing_fields.append("Salesperson Name")
        if not site_location.strip():
            missing_fields.append("Site Location")
        if not application_overview.strip():
            missing_fields.append("Application Overview")

        # System Config
        if not materials:
            missing_fields.append("Materials to Sort")
        if conveyor_included == "Yes" and not conveyor_size.strip():
            missing_fields.append("Conveyor Size")
        if not belt_speed.strip():
            missing_fields.append("Belt Speed")
        if not pick_rate.strip():
            missing_fields.append("Pick Rate")
        if not robot_type:
            missing_fields.append("Robot Type")
        if not gripper_type:
            missing_fields.append("Gripper Type")

        # Technical Specs
        if max_object_weight == 0.0:
            missing_fields.append("Max Object Weight")
        if robot_bases == 0:
            missing_fields.append("Number of Robot Bases")
        if not vision_system:
            missing_fields.append("Vision System")
        if input_power_kva == 0.0:
            missing_fields.append("Input Power")
        if avg_consumption_kw == 0.0:
            missing_fields.append("Average Power Consumption")
        if air_consumption_lpm == 0:
            missing_fields.append("Air Consumption")

        # Shipping & Timeline
        if not shipping_distance.strip():
            missing_fields.append("Shipping Distance")
        for step, duration in timeline.items():
            if not duration.strip():
                missing_fields.append(f"Timeline: {step}")

        if missing_fields:
            st.error("Please fill in all required fields:\n- " + "\n- ".join(missing_fields))
            st.stop()

        # --- SAFE TO EXECUTE BELOW THIS LINE ---

        # Calculate pricing
        def calculate_price_breakdown(inputs):
            breakdown = []

            breakdown.append({
                "Component": "Robot Arm",
                "Description": inputs["robot_type"],
                "Unit Price": PRICING["robot_arm"],
                "Qty": inputs["robot_arms"],
                "Subtotal": PRICING["robot_arm"] * inputs["robot_arms"]
            })

            breakdown.append({
                "Component": "Grippers",
                "Description": ", ".join(inputs["gripper_type"]),
                "Unit Price": PRICING["gripper"],
                "Qty": len(inputs["gripper_type"]),
                "Subtotal": PRICING["gripper"] * len(inputs["gripper_type"])
            })

            if inputs["conveyor_included"] == "Yes":
                breakdown.append({
                    "Component": "Conveyor",
                    "Description": f"{inputs['conveyor_size']} inch belt",
                    "Unit Price": PRICING["conveyor"],
                    "Qty": 1,
                    "Subtotal": PRICING["conveyor"]
                })

            breakdown.append({
                "Component": "Hypervision Scanners",
                "Description": "2D + 3D Scanning",
                "Unit Price": PRICING["hypervision_scanner"],
                "Qty": inputs["hypervision_scanners"],
                "Subtotal": PRICING["hypervision_scanner"] * inputs["hypervision_scanners"]
            })

            if inputs["ai_training"]:
                breakdown.append({
                    "Component": "AI Custom Training",
                    "Description": "Trash / PCBs / UBCs",
                    "Unit Price": PRICING["ai_training"],
                    "Qty": 1,
                    "Subtotal": PRICING["ai_training"]
                })

            if inputs["software_license"]:
                breakdown.append({
                    "Component": "Software License",
                    "Description": "Perpetual",
                    "Unit Price": PRICING["software_license"],
                    "Qty": 1,
                    "Subtotal": PRICING["software_license"]
                })

            if inputs["fat"]:
                breakdown.append({
                    "Component": "Factory Pre-Assembly (FAT)",
                    "Description": "Pre-assembly and testing",
                    "Unit Price": PRICING["fat"],
                    "Qty": 1,
                    "Subtotal": PRICING["fat"]
                })

            if inputs["try_and_buy"]:
                breakdown.append({
                    "Component": "Try & Buy Second Arm",
                    "Description": "Deferred Payment",
                    "Unit Price": PRICING["try_and_buy_arm"],
                    "Qty": 1,
                    "Subtotal": PRICING["try_and_buy_arm"]
                })

            if inputs["installation_by_wr"] == "Yes":
                breakdown.append({
                    "Component": "Installation by WR",
                    "Description": "Full on-site implementation",
                    "Unit Price": PRICING["installation"],
                    "Qty": 1,
                    "Subtotal": PRICING["installation"]
                })

            try:
                dist = float(inputs["shipping_distance"])
                shipping_cost = dist * PRICING["shipping_per_km"]
                breakdown.append({
                    "Component": "Shipping",
                    "Description": f"{dist} km at ${PRICING['shipping_per_km']}/km",
                    "Unit Price": PRICING["shipping_per_km"],
                    "Qty": dist,
                    "Subtotal": shipping_cost
                })
            except:
                pass

            mod_key = f"modification_{inputs['modification_scale'].lower()}"
            breakdown.append({
                "Component": "Modification Scope",
                "Description": inputs["modification_scale"],
                "Unit Price": PRICING[mod_key],
                "Qty": 1,
                "Subtotal": PRICING[mod_key]
            })

            if inputs["security_perimeter"]:
                breakdown.append({
                    "Component": "Security Perimeter",
                    "Description": "Robot safety fencing",
                    "Unit Price": PRICING["security_perimeter"],
                    "Qty": 1,
                    "Subtotal": PRICING["security_perimeter"]
                })

            if inputs["warranty"]:
                breakdown.append({
                    "Component": "Warranty (1 year)",
                    "Description": "Parts + labor coverage",
                    "Unit Price": PRICING["warranty"],
                    "Qty": 1,
                    "Subtotal": PRICING["warranty"]
                })

            if inputs["pe_stamp"]:
                breakdown.append({
                    "Component": "PE Stamp",
                    "Description": "Professional engineer review",
                    "Unit Price": PRICING["pe_stamp"],
                    "Qty": 1,
                    "Subtotal": PRICING["pe_stamp"]
                })

            if inputs["sat"]:
                breakdown.append({
                    "Component": "Site Acceptance Test (SAT)",
                    "Description": "Final performance check",
                    "Unit Price": PRICING["sat"],
                    "Qty": 1,
                    "Subtotal": PRICING["sat"]
                })

            return breakdown

        inputs = {
            "robot_arms": robot_arms,
            "gripper_type": gripper_type,
            "conveyor_included": conveyor_included,
            "conveyor_size": conveyor_size,
            "hypervision_scanners": hypervision_scanners,
            "ai_training": ai_training,
            "software_license": software_license,
            "fat": fat,
            "try_and_buy": try_and_buy,
            "installation_by_wr": installation_by_wr,
            "robot_type": robot_type,
            "shipping_distance": shipping_distance,
            "modification_scale": modification_scale,
            "security_perimeter": security_perimeter,
            "warranty": warranty,
            "pe_stamp": pe_stamp,
            "sat": sat
        }

        df = pd.DataFrame(calculate_price_breakdown(inputs))
        multiplier = CURRENCY_CONVERSION.get(currency, 1.0)
        df["Unit Price"] *= multiplier
        df["Subtotal"] *= multiplier

        total = df["Subtotal"].sum()
        st.dataframe(df.style.format({"Unit Price": "${:,.0f}", "Subtotal": "${:,.0f}"}))
        st.markdown(f"### **Total Estimated Price: {currency} {total:,.0f}**")

        # --- DOCX Generation ---
        doc = DocxTemplate("template_practice.docx")
        layout_image = InlineImage(doc, f"layout_{robot_arms}_arms.png", width=Mm(100), height=Mm(80))
        layout_overview_image = InlineImage(doc, f"overview_{robot_arms}_arms.png", width=Mm(150), height=Mm(80))
        robot_model_image = InlineImage(doc, "fanuc_m20id25.png", width=Mm(100), height=Mm(80))

        if gripper_type:
            base_name = gripper_type[0].lower().replace(" ", "_")
            gripper_filename = f"gripper_{base_name}.png"
        else:
            gripper_filename = "gripper_default.png"

        if not os.path.exists(gripper_filename):
            gripper_filename = "gripper_default.png"

        gripper_image = InlineImage(doc, gripper_filename, width=Mm(100), height=Mm(80))

        context = {
            "value_proposition": value_proposition,
            "application_overview": application_overview,
            "client_name": client_name,
            "client_company": client_company,
            "quote_date": quote_date.strftime("%B %d, %Y"),
            "site_location": site_location,
            "robot_type": robot_type,
            "robot_arms": robot_arms,
            "materials": ", ".join(materials),
            "belt_speed": belt_speed,
            "pick_rate": pick_rate,
            "max_object_weight": max_object_weight,
            "robot_bases": robot_bases,
            "vision_system": ", ".join(vision_system),
            "input_power_kva": input_power_kva,
            "avg_consumption_kw": avg_consumption_kw,
            "air_consumption_lpm": air_consumption_lpm,
            "timeline": timeline,
            "total_price": f"{currency} {total:,.0f}",
            "layout_image": layout_image,
            "gripper_type": ", ".join(gripper_type),
            "additional_arm_price": f"{currency} {PRICING['try_and_buy_arm'] * multiplier:,.0f}",
            "gripper_image": gripper_image,
            "layout_overview_image": layout_overview_image,
            "robot_model_image": robot_model_image
        }

        file_name = f"{client_name}_Quote_{quote_date.strftime('%Y%m%d')}.docx"
        doc.render(context)
        doc.save(file_name)

        st.success("✅ Quote generated successfully!")

        with open(file_name, "rb") as f:
            st.download_button(
                label="📄 Download Quote DOCX",
                data=f,
                file_name=file_name,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

st.markdown("""
    <div class="footer" style='
        position: fixed;
        bottom: 0;
        width: 100%;
        background-color: #1A1A1A;
        color: white;
        text-align: center;
        padding: 5px;
        font-size: 12px;
    '>
        © 2025 Waste Robotics | Internal Tool
    </div>
""", unsafe_allow_html=True)



