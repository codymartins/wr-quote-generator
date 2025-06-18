# import streamlit as st
# import pandas as pd

# # Currency Conversion
# CURRENCY_CONVERSION = {
#     "USD": 1.0,
#     "CAD": 1.36,
#     "EUR": 0.92
# }

# # Base Pricing Estimates (in USD)
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
#     "infrastructure_electrical": 4000,
#     "infrastructure_air": 3500,
#     "infrastructure_internet": 1500,
#     "security_perimeter": 2500,
#     "warranty": 8000,
#     "pe_stamp": 2000,
#     "sat": 6000
# }

# def calculate_price_breakdown(inputs):
#     breakdown = []

#     breakdown.append({
#         "Component": "Robot Arm",
#         "Description": inputs["robot_type"],
#         "Unit Price": PRICING["robot_arm"],
#         "Qty": inputs["robot_arms"],
#         "Subtotal": PRICING["robot_arm"] * inputs["robot_arms"]
#     })

#     breakdown.append({
#         "Component": "Grippers",
#         "Description": ", ".join(inputs["gripper_type"]),
#         "Unit Price": PRICING["gripper"],
#         "Qty": len(inputs["gripper_type"]),
#         "Subtotal": PRICING["gripper"] * len(inputs["gripper_type"])
#     })

#     if inputs["conveyor_included"] == "Yes":
#         breakdown.append({
#             "Component": "Conveyor",
#             "Description": f"{inputs['conveyor_size']} inch belt",
#             "Unit Price": PRICING["conveyor"],
#             "Qty": 1,
#             "Subtotal": PRICING["conveyor"]
#         })

#     breakdown.append({
#         "Component": "Hypervision Scanners",
#         "Description": "2D + 3D Scanning",
#         "Unit Price": PRICING["hypervision_scanner"],
#         "Qty": inputs["hypervision_scanners"],
#         "Subtotal": PRICING["hypervision_scanner"] * inputs["hypervision_scanners"]
#     })

#     if inputs["ai_training"]:
#         breakdown.append({
#             "Component": "AI Custom Training",
#             "Description": "Trash / PCBs / UBCs",
#             "Unit Price": PRICING["ai_training"],
#             "Qty": 1,
#             "Subtotal": PRICING["ai_training"]
#         })

#     if inputs["software_license"]:
#         breakdown.append({
#             "Component": "Software License",
#             "Description": "Perpetual",
#             "Unit Price": PRICING["software_license"],
#             "Qty": 1,
#             "Subtotal": PRICING["software_license"]
#         })

#     if inputs["fat"]:
#         breakdown.append({
#             "Component": "Factory Pre-Assembly (FAT)",
#             "Description": "Pre-assembly and testing",
#             "Unit Price": PRICING["fat"],
#             "Qty": 1,
#             "Subtotal": PRICING["fat"]
#         })

#     if inputs["try_and_buy"]:
#         breakdown.append({
#             "Component": "Try & Buy Second Arm",
#             "Description": "Deferred Payment",
#             "Unit Price": PRICING["try_and_buy_arm"],
#             "Qty": 1,
#             "Subtotal": PRICING["try_and_buy_arm"]
#         })

#     if inputs["installation_by_wr"] == "Yes":
#         breakdown.append({
#             "Component": "Installation by WR",
#             "Description": "Full on-site implementation",
#             "Unit Price": PRICING["installation"],
#             "Qty": 1,
#             "Subtotal": PRICING["installation"]
#         })

#     try:
#         dist = float(inputs["shipping_distance"])
#         shipping_cost = dist * PRICING["shipping_per_km"]
#         breakdown.append({
#             "Component": "Shipping",
#             "Description": f"{dist} km at ${PRICING['shipping_per_km']}/km",
#             "Unit Price": PRICING["shipping_per_km"],
#             "Qty": dist,
#             "Subtotal": shipping_cost
#         })
#     except:
#         pass

#     mod_key = f"modification_{inputs['modification_scale'].lower()}"
#     breakdown.append({
#         "Component": "Modification Scope",
#         "Description": inputs["modification_scale"],
#         "Unit Price": PRICING[mod_key],
#         "Qty": 1,
#         "Subtotal": PRICING[mod_key]
#     })

#     if "Electrical Power" in inputs["robotic_cell_requirements"]:
#         breakdown.append({
#             "Component": "Infrastructure - Electrical",
#             "Description": "Power hookup by WR",
#             "Unit Price": PRICING["infrastructure_electrical"],
#             "Qty": 1,
#             "Subtotal": PRICING["infrastructure_electrical"]
#         })

#     if "Compressed Air Supply" in inputs["robotic_cell_requirements"]:
#         breakdown.append({
#             "Component": "Infrastructure - Compressed Air",
#             "Description": "Air hookup by WR",
#             "Unit Price": PRICING["infrastructure_air"],
#             "Qty": 1,
#             "Subtotal": PRICING["infrastructure_air"]
#         })

#     if "Internet Connection" in inputs["robotic_cell_requirements"]:
#         breakdown.append({
#             "Component": "Infrastructure - Internet",
#             "Description": "Network setup by WR",
#             "Unit Price": PRICING["infrastructure_internet"],
#             "Qty": 1,
#             "Subtotal": PRICING["infrastructure_internet"]
#         })

#     if inputs["security_perimeter"]:
#         breakdown.append({
#             "Component": "Security Perimeter",
#             "Description": "Robot safety fencing",
#             "Unit Price": PRICING["security_perimeter"],
#             "Qty": 1,
#             "Subtotal": PRICING["security_perimeter"]
#         })

#     if inputs["warranty"]:
#         breakdown.append({
#             "Component": "Warranty (1 year)",
#             "Description": "Parts + labor coverage",
#             "Unit Price": PRICING["warranty"],
#             "Qty": 1,
#             "Subtotal": PRICING["warranty"]
#         })

#     if inputs["pe_stamp"]:
#         breakdown.append({
#             "Component": "PE Stamp",
#             "Description": "Professional engineer review",
#             "Unit Price": PRICING["pe_stamp"],
#             "Qty": 1,
#             "Subtotal": PRICING["pe_stamp"]
#         })

#     if inputs["sat"]:
#         breakdown.append({
#             "Component": "Site Acceptance Test (SAT)",
#             "Description": "Final performance check",
#             "Unit Price": PRICING["sat"],
#             "Qty": 1,
#             "Subtotal": PRICING["sat"]
#         })

#     return breakdown


# st.title("Waste Robotics Quote Generator")

# # Section 1: Basic Info
# st.header("Client and Quote Details")
# quote_date = st.date_input("Quote Date")
# quote_name = st.text_input("Quote Name")
# client_name = st.text_input("Client Name")
# salesman_name = st.text_input("Salesperson Name")
# site_location = st.text_input("Site Location")
# currency = st.selectbox("Currency", ["USD", "CAD", "EUR"])

# # Section 2: System Details
# st.header("System Configuration")
# materials = st.multiselect("Materials to Sort", ["PCBs", "UBCs", "Trash", "Other"])
# try_and_buy = st.checkbox("Include Try & Buy Option?")
# conveyor_included = st.selectbox("Conveyor Provided?", ["Yes", "No"])
# conveyor_size = st.text_input("Conveyor Size (inches)", disabled=(conveyor_included == "No"))
# belt_speed = st.text_input("Belt Speed (e.g. 80 ft/min)")
# pick_rate = st.text_input("Pick Rate (e.g. 35 picks/minute)")

# modification_scale = st.selectbox("Modification Scope", ["Small", "Medium", "Large"])
# hypervision_scanners = st.number_input("Number of Hypervision Scanners", min_value=0, value=1)

# robot_type = st.selectbox("Robot Type", ["Fanuc LR-Mate 200ID", "Other"])
# robot_arms = st.slider("Number of Robot Arms", 1, 3, 2)
# gripper_type = st.multiselect("Gripper Type", ["VentuR (suction)", "PinchR", "Other"])

# # Section 3: Infrastructure Requirements
# st.header("Robotic Cell Functioning Requirements")
# robotic_cell_requirements = st.multiselect(
#     "Client-Provided Infrastructure",
#     ["Electrical Power", "Compressed Air Supply", "Internet Connection"]
# )

# shipping_distance = st.text_input("Estimated Shipping Distance (miles or km)")

# # Section 4: Timeline
# st.header("Project Timeline Estimates")
# timeline = {
#     "Order Confirmation / Project Kickoff": st.text_input("Order Confirmation / Project Kickoff Duration"),
#     "Detailed Engineering": st.text_input("Detailed Engineering Duration"),
#     "Engineering Review": st.text_input("Engineering Review Duration"),
#     "Procurement and Fabrication": st.text_input("Procurement and Fabrication Duration"),
#     "FAT and Shipping": st.text_input("FAT and Shipping Duration"),
#     "Retrofit and Installation": st.text_input("Retrofit and Installation Duration"),
#     "Commissioning and SAT": st.text_input("Commissioning and SAT Duration")
# }

# # Section 5: Inclusions
# st.header("Inclusions and Add-ons")
# security_perimeter = st.checkbox("Include Security Perimeter?")
# software_license = st.checkbox("Include Perpetual Software License?")
# ai_training = st.checkbox("Include AI Custom Training?")
# warranty = st.checkbox("Include Warranty?")
# fat = st.checkbox("Include Factory Pre-Assembly and Testing (FAT)?")
# pe_stamp = st.checkbox("Include PE Stamp?")
# sat = st.checkbox("Include Site Acceptance Test (SAT)?")

# # Section 6: Notes
# st.header("Other Considerations")
# installation_by_wr = st.selectbox("Will WR handle site modifications and installation?", ["Yes", "No", "Partially"])

# # Submit Button
# if st.button("Generate Quote"):
#         inputs = {
#         "robot_arms": robot_arms,
#         "gripper_type": gripper_type,
#         "conveyor_included": conveyor_included,
#         "conveyor_size": conveyor_size,
#         "hypervision_scanners": hypervision_scanners,
#         "ai_training": ai_training,
#         "software_license": software_license,
#         "fat": fat,
#         "try_and_buy": try_and_buy,
#         "installation_by_wr": installation_by_wr,
#         "robot_type": robot_type,
#         "shipping_distance": shipping_distance,
#         "modification_scale": modification_scale,
#         "robotic_cell_requirements": robotic_cell_requirements,
#         "security_perimeter": security_perimeter,
#         "warranty": warranty,
#         "pe_stamp": pe_stamp,
#         "sat": sat
#     }

# breakdown = calculate_price_breakdown(inputs)
# df = pd.DataFrame(breakdown)

# currency_multiplier = CURRENCY_CONVERSION.get(currency, 1.0)
# df["Unit Price"] *= currency_multiplier
# df["Subtotal"] *= currency_multiplier

# st.subheader("💲 Quote Price Breakdown")
# st.dataframe(df.style.format({"Unit Price": "${:,.0f}", "Subtotal": "${:,.0f}"}))

# total = df["Subtotal"].sum()
# st.markdown(f"### **Total Estimated Price: {currency} {total:,.0f}**")

# csv = df.to_csv(index=False).encode('utf-8')

# st.download_button(
#     label="📥 Download Price Breakdown as CSV",
#     data=csv,
#     file_name=f"{quote_name or 'WR_Quote'}_breakdown.csv",
#     mime="text/csv"
# )
# from docxtpl import DocxTemplate, InlineImage
# from docx.shared import Mm
# import pythoncom
# from docx2pdf import convert
# import os

# # Step 1: Create one single doc object
# doc = DocxTemplate("template_practice.docx")

# # Step 2: Select and insert the correct image
# image_path = f"layout_{robot_arms}_arms.png"

# # Optional: fallback image if file not found
# if not os.path.exists(image_path):
#     image_path = "layout_default.png"

# layout_image = InlineImage(doc, image_path, width=Mm(150))

# # Step 3: Build context with image
# context = {
#     "client_name": client_name,
#     "quote_date": quote_date.strftime("%B %d, %Y"),
#     "site_location": site_location,
#     "robot_type": robot_type,
#     "robot_arms": robot_arms,
#     "materials": ", ".join(materials),
#     "total_price": f"{currency} {total:,.0f}",
#     "layout_image": layout_image
# }

# # Step 4: Render and save as DOCX
# doc.render(context)
# doc.save("generated_quote.docx")

# # ✅ Step 5: Convert to PDF safely using COM initialization
# def safe_convert_docx_to_pdf(input_path, output_path):
#     pythoncom.CoInitialize()
#     convert(input_path, output_path)
#     pythoncom.CoUninitialize()

# safe_convert_docx_to_pdf("generated_quote.docx", "generated_quote.pdf")

# # Step 6: Let the user download the PDF
# with open("generated_quote.pdf", "rb") as f:
#     st.download_button(
#         label="📄 Download Quote PDF",
#         data=f,
#         file_name="WR_Practice_Quote.pdf",
#         mime="application/pdf"
#     )

import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
from docx2pdf import convert
import pythoncom
import os

# --- Constants ---
CURRENCY_CONVERSION = {"USD": 1.0, "CAD": 1.36, "EUR": 0.92}
PRICING = {
    "robot_arm": 45000,
    "gripper": 6000,
    "conveyor": 8000,
    "hypervision_scanner": 12000,
    "ai_training": 15000,
    "software_license": 5000,
    "fat": 4000,
    "try_and_buy_arm": 90000,
    "installation": 30000,
    "shipping_per_km": 3.5,
    "modification_small": 5000,
    "modification_medium": 15000,
    "modification_large": 30000,
    "security_perimeter": 2500,
    "warranty": 8000,
    "pe_stamp": 2000,
    "sat": 6000
}

# --- UI ---
st.title("Waste Robotics Quote Generator")

# Section 1: Client + Project Info
st.header("Client and Proposal Details")
quote_date = st.date_input("Quote Date")
value_proposition = st.text_input("Value Proposition (Main Proposal Title)")
client_name = st.text_input("Client Name")
client_company = st.text_input("Client Company Name")
salesman_name = st.text_input("Salesperson Name")
site_location = st.text_input("Site Location")
currency = st.selectbox("Currency", ["USD", "CAD", "EUR"])

# Section 2: Application Overview
st.header("Application Overview")
application_overview = st.text_area("Brief Summary of the Application")

# Section 3: System Configuration
st.header("System Configuration")
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

# Section 4: Technical Details
st.header("Technical Details")
max_object_weight = st.number_input("Maximum Object Weight per Robot (kg)", min_value=0.0)
robot_bases = st.number_input("Number of Robot Bases", min_value=0)
vision_system = st.multiselect("Robot Vision System", ["RGB Camera", "3D Camera"])
input_power_kva = st.number_input("Input Power (kVA)", min_value=0.0)
avg_consumption_kw = st.number_input("Average Power Consumption (kW)", min_value=0.0)
air_consumption_lpm = st.number_input("Total Air Consumption (L/min)", min_value=0)

# Section 5: Shipping
shipping_distance = st.text_input("Estimated Shipping Distance (miles or km)")

# Section 6: Timeline
st.header("Project Timeline Estimates")
timeline = {
    "Order Confirmation / Project Kickoff": st.text_input("Order Confirmation / Project Kickoff Duration"),
    "Detailed Engineering": st.text_input("Detailed Engineering Duration"),
    "Engineering Review": st.text_input("Engineering Review Duration"),
    "Procurement and Fabrication": st.text_input("Procurement and Fabrication Duration"),
    "FAT and Shipping": st.text_input("FAT and Shipping Duration"),
    "Retrofit and Installation": st.text_input("Retrofit and Installation Duration"),
    "Commissioning and SAT": st.text_input("Commissioning and SAT Duration")
}

# Section 7: Add-ons
st.header("Inclusions and Add-ons")
security_perimeter = st.checkbox("Include Security Perimeter?")
software_license = st.checkbox("Include Perpetual Software License?")
ai_training = st.checkbox("Include AI Custom Training?")
warranty = st.checkbox("Include Warranty?")
fat = st.checkbox("Include Factory Pre-Assembly and Testing (FAT)?")
pe_stamp = st.checkbox("Include PE Stamp?")
sat = st.checkbox("Include Site Acceptance Test (SAT)?")
installation_by_wr = st.selectbox("Will WR handle site modifications and installation?", ["Yes", "No", "Partially"])

# --- Quote Generation ---
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

# --- Render Quote ---
if st.button("Generate Quote"):
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

    breakdown = calculate_price_breakdown(inputs)
    df = pd.DataFrame(breakdown)
    currency_multiplier = CURRENCY_CONVERSION.get(currency, 1.0)
    df["Unit Price"] *= currency_multiplier
    df["Subtotal"] *= currency_multiplier

    st.subheader("💲 Quote Price Breakdown")
    st.dataframe(df.style.format({"Unit Price": "${:,.0f}", "Subtotal": "${:,.0f}"}))
    total = df["Subtotal"].sum()
    st.markdown(f"### **Total Estimated Price: {currency} {total:,.0f}**")

    # --- Document Generation ---
    doc = DocxTemplate("template_practice.docx")
    layout_image = InlineImage(doc, f"layout_{robot_arms}_arms.png", width=Mm(100), height=Mm(80))
    layout_overview_image = InlineImage(doc, f"overview_{robot_arms}_arms.png", width=Mm(150), height=Mm(80))
    robot_model_image = InlineImage(doc, "fanuc_m20id25.png", width=Mm(100), height=Mm(80))
    # Build gripper image path from first selected gripper
    try:
        gripper_filename = f"gripper_{gripper_type[0].lower().replace(' ', '_')}.png"
    except IndexError:
        gripper_filename = "gripper_default.png"  # fallback if nothing selected

    gripper_image = InlineImage(doc, gripper_filename, width=Mm(100), height=Mm(80))

    context = {
        "value_proposition": value_proposition,
        "application_overview": application_overview,
        "client_name": client_name,
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
        "layout_overview_image": layout_overview_image,
        "robot_model_image": robot_model_image,
        "client_company": client_company,
        "pick_rate": pick_rate,
        "gripper_type": ", ".join(gripper_type),
        "site_location": site_location,
        "additional_arm_price": f"{currency} {PRICING['try_and_buy_arm'] * CURRENCY_CONVERSION.get(currency, 1.0):,.0f}",
        "gripper_image": gripper_image

    }

    doc.render(context)
    doc.save("generated_quote.docx")

    # def safe_convert_docx_to_pdf(input_path, output_path):
    #     pythoncom.CoInitialize()
    #     convert(input_path, output_path)
    #     pythoncom.CoUninitialize()

    # safe_convert_docx_to_pdf("generated_quote.docx", "generated_quote.pdf")

    # with open("generated_quote.pdf", "rb") as f:
    #     st.download_button(
    #         label="📄 Download Quote PDF",
    #         data=f,
    #         file_name="WR_Practice_Quote.pdf",
    #         mime="application/pdf"
    #     )

# Step 4: Render and save the .docx
doc.render(context)
file_name = f"{client_name}_Quote_{quote_date.strftime('%Y%m%d')}.docx"
doc.save(file_name)

# Step 5: Display success message and download button
st.success("✅ Quote generated successfully!")

with open(file_name, "rb") as f:
    st.download_button(
        label="📄 Download Quote DOCX",
        data=f,
        file_name=file_name,
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )