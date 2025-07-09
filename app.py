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
import matplotlib.pyplot as plt
import tempfile
from io import BytesIO

# --- WR Branding Setup ---
logo = Image.open("logoWasteRobotics(1).png")  # Make sure this file is in the same directory
col1, col2 = st.columns([1, 6])
with col1:
    st.image("logoWasteRobotics(1).png", width=80)
with col2:
    st.markdown("<h1 style='color: white;'>Waste Robotics Quote Generator</h1>", unsafe_allow_html=True)
    st.markdown("<p style='color: #EF3A2D; font-style: italic;'>Smarter Sorting with Robotics</p>", unsafe_allow_html=True)

def save_df_as_image(df, currency="CAD"):
    df = df.copy()
    df["Unit Price"] = df["Unit Price"].map(lambda x: f"{currency} {x:,.0f}")
    df["Subtotal"] = df["Subtotal"].map(lambda x: f"{currency} {x:,.0f}")

    fig, ax = plt.subplots(figsize=(10, len(df) * 0.5 + 1))
    ax.axis("off")

    table = ax.table(
        cellText=df.values,
        colLabels=df.columns,
        cellLoc='left',
        loc='center'
    )
    table.auto_set_font_size(False)
    table.set_fontsize(10)

    for i in range(len(df.columns)):
        table[0, i].set_facecolor("#ef3a2d")
        table[0, i].set_text_props(weight="bold", color="white")

    for row_idx in range(1, len(df) + 1):
        color = "#f2f2f2" if row_idx % 2 == 0 else "#ffffff"
        for col_idx in range(len(df.columns)):
            table[row_idx, col_idx].set_facecolor(color)

    fig.tight_layout()

    buf = BytesIO()
    plt.savefig(buf, format="png", dpi=300, bbox_inches="tight")
    plt.close(fig)
    buf.seek(0)
    return buf


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

# Convert it to a dictionary for easy lookup (now using price_cad, always as float)
PRICING = {
    row["item"]: float(str(row["price_cad"]).replace(",", ""))
    for _, row in pricing_df.iterrows()
}

# --- Constants ---
# All prices in CSV are in CAD, so CAD is the base currency
CURRENCY_CONVERSION = {"CAD": 1.0, "USD": 0.74, "EUR": 0.68}
# PRICING = {
#     "robot_arm": 45000,
#     "gripper": 6000,
#     "conveyor": 8000,
#     "hypervision_scanner": 12000,
#     "ai_training": 15000,
#     "software_license": 5000,l
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
    shipping_method = st.selectbox("Shipping Method", ["Truck", "Boat"], help="Select the shipping method for delivery.")
    if shipping_method == "Truck":
        num_trucks_or_containers = st.number_input("Number of Trucks", min_value=1, value=1, step=1)
    else:
        num_trucks_or_containers = st.number_input("Number of Containers (Boat)", min_value=1, value=1, step=1)
    currency = st.selectbox("Currency", ["USD", "CAD", "EUR"])
    application_overview = st.text_area("Brief Summary of the Application")


with tab2:
    st.header("System Configuration")
    st.progress(40, text="Step 2 of 5")
    materials = st.multiselect("Materials to Sort", ["PCBs", "UBCs", "Trash", "Other"])
    try_and_buy = st.checkbox("Include Try & Buy Option?")
    belt_speed = st.text_input("Belt Speed (e.g. 80 ft/min)")
    pick_rate = st.text_input("Pick Rate (e.g. 35 picks/minute)")
    # Robot Arms (type and quantity)
    robot_types_list = ["Fanuc LR-Mate", "FanucLr10iA", "Fanuc Delta DR3", "Fanuc M10", "Fanuc M20", "Fanuc M710"]
    selected_robot_types = st.multiselect("Robot Arm Types", robot_types_list)
    robot_type = {}
    for rtype in selected_robot_types:
        qty = st.number_input(f"Quantity of {rtype}", min_value=0, value=1, key=f"qty_robot_{rtype}")
        if qty > 0:
            robot_type[rtype] = qty

    # Robot Bases (type and quantity)
    base_types = ["LrMate/Lr10ia", "Delta DR3", "M-10, M-20, M-710"]
    selected_bases = st.multiselect("Robot Base Types", base_types)
    robot_bases = {}
    for base in selected_bases:
        qty = st.number_input(f"Quantity of {base}", min_value=0, value=1, key=f"qty_base_{base}")
        if qty > 0:
            robot_bases[base] = qty

    # Grippers (type and quantity)
    gripper_types_list = ["VentuR", "BagR", "BagR CO", "PinchR Lr & M10", "MonstR", "DagR"]
    selected_grippers = st.multiselect("Gripper Types", gripper_types_list)
    gripper_type = {}
    for gtype in selected_grippers:
        qty = st.number_input(f"Quantity of {gtype}", min_value=0, value=1, key=f"qty_gripper_{gtype}")
        if qty > 0:
            gripper_type[gtype] = qty

    # Consistency check
    total_arms = sum(robot_type.values()) if robot_type else 0
    total_bases = sum(robot_bases.values()) if robot_bases else 0
    total_grippers = sum(gripper_type.values()) if gripper_type else 0
    if total_arms != total_bases or total_arms != total_grippers:
        st.warning(f"⚠️ The total number of robot arms ({total_arms}), robot bases ({total_bases}), and grippers ({total_grippers}) should be the same for a valid configuration.")


with tab3:
    st.header("Technical Specs")
    st.progress(60, text="Step 3 of 5")
    max_object_weight = st.number_input("Maximum Object Weight per Robot (kg)", min_value=0.0)
    # Disposition prompt
    disposition = st.selectbox("Disposition", ["FTF", "IL", "N/A", "QCX"])
    # VRS Model prompt
    vrs_model = st.selectbox("VRS Model", ["900", "1200", "1600", "1800"])
    # Vision System (type and quantity)
    vision_types_list = ["DeepVision System", "HyperVision System"]
    selected_vision_types = st.multiselect("Robot Vision System", vision_types_list)
    vision_system = {}
    for vtype in selected_vision_types:
        qty = st.number_input(f"Quantity of {vtype}", min_value=0, value=1, key=f"qty_vision_{vtype}")
        if qty > 0:
            vision_system[vtype] = qty
    input_power_kva = st.number_input("Input Power (kVA)", min_value=0.0)
    avg_consumption_kw = st.number_input("Average Power Consumption (kW)", min_value=0.0)
    air_consumption_lpm = st.number_input("Total Air Consumption (L/min)", min_value=0)

with tab4:
    st.header("Shipping & Timeline")
    st.progress(80, text="Step 4 of 5")
    # shipping_distance = st.text_input("Estimated Shipping Distance (miles or km)")
    # Shipping distance is now replaced by method/count logic
    order_confirmation_project_kickoff = st.text_input("Order Confirmation / Project Kickoff Duration")
    detailed_engineering = st.text_input("Detailed Engineering Duration")
    engineering_review = st.text_input("Engineering Review Duration")
    procurement_fabrication = st.text_input("Procurement and Fabrication Duration")
    fat_shipping = st.text_input("FAT and Shipping Duration")
    retrofit_installation = st.text_input("Retrofit and Installation Duration")
    commissioning_and_SAT = st.text_input("Commissioning and SAT Duration")

with tab5:

    st.header("Inclusions & Final Quote")
    st.progress(100, text="Step 5 of 5")

    colA, colB = st.columns(2)
    with colA:
        safety_fencing = st.checkbox("Include Safety Fencing?")
        conveyor_var_speed_license = st.checkbox("Include Conveyor Variable Speed License?")
        custom_ai_training = st.checkbox("Include Custom AI Training?")
        robot_validator_license = st.checkbox("Include Robot Validator License?")
        greyparrot_monitoring_unit = st.checkbox("Include GreyParrot Monitoring Unit?")
        installation_supervision = st.checkbox("Include Installation Supervision?")
    with colB:
        additional_sorting_recipes = st.checkbox("Include Additional Sorting Recipes?")
        sat_to_cfa = st.checkbox("Include SAT to CFA?")
        engineering_and_documentation = st.checkbox("Include Engineering & Documentation?")
        online_commissioning = st.checkbox("Include Online Commissioning?")
        installation_commissioning_training = st.checkbox("Include Installation, Commissioning & Training?")
        lips2_support = st.checkbox("Include LIPS2 Support?")
        warranty_option = st.selectbox(
            "Warranty Option",
            ["None", "1 Year (Standard)", "Extended"]
        )

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
        # Shipping method/count validation
        if not site_location.strip():
            missing_fields.append("Site Location")
        if not shipping_method:
            missing_fields.append("Shipping Method")
        if not num_trucks_or_containers or num_trucks_or_containers < 1:
            missing_fields.append("Number of Trucks/Containers")
        if not order_confirmation_project_kickoff.strip():
            missing_fields.append("Order Confirmation / Project Kickoff Duration")
        if not detailed_engineering.strip():
            missing_fields.append("Detailed Engineering Duration")
        if not engineering_review.strip():
            missing_fields.append("Engineering Review Duration")
        if not procurement_fabrication.strip():
            missing_fields.append("Procurement and Fabrication Duration")
        if not fat_shipping.strip():
            missing_fields.append("FAT and Shipping Duration")
        if not retrofit_installation.strip():
            missing_fields.append("Retrofit and Installation Duration")
        if not commissioning_and_SAT.strip():
            missing_fields.append("Commissioning and SAT Duration")

        if missing_fields:
            st.error("Please fill in all required fields:\n- " + "\n- ".join(missing_fields))
            st.stop()

        # --- SAFE TO EXECUTE BELOW THIS LINE ---

        # Calculate pricing
        def calculate_price_breakdown(inputs):
            # Only include items that are present in the current PRICING dict (from pricing.csv)
            breakdown = []
            if "Conveyor_Variable_Speed_License" in PRICING and inputs.get("conveyor_var_speed_license"):
                breakdown.append({
                    "Component": "Conveyor Variable Speed License",
                    "Description": "Conveyor Variable Speed License",
                    "Unit Price": PRICING.get("Conveyor_Variable_Speed_License", 0),
                    "Qty": 1,
                    "Subtotal": PRICING.get("Conveyor_Variable_Speed_License", 0)
                })
            if "custom_ai_training" in PRICING and inputs.get("custom_ai_training"):
                breakdown.append({
                    "Component": "Custom AI Training",
                    "Description": "Custom AI Training",
                    "Unit Price": PRICING.get("custom_ai_training", 0),
                    "Qty": 1,
                    "Subtotal": PRICING.get("custom_ai_training", 0)
                })
            if "robot_validator_license" in PRICING and inputs.get("robot_validator_license"):
                breakdown.append({
                    "Component": "Robot Validator License",
                    "Description": "Robot Validator License",
                    "Unit Price": PRICING.get("robot_validator_license", 0),
                    "Qty": 1,
                    "Subtotal": PRICING.get("robot_validator_license", 0)
                })
            if "GreyParrot_Monitoring_Unit" in PRICING and inputs.get("greyparrot_monitoring_unit"):
                breakdown.append({
                    "Component": "GreyParrot Monitoring Unit",
                    "Description": "GreyParrot Monitoring Unit",
                    "Unit Price": PRICING.get("GreyParrot_Monitoring_Unit", 0),
                    "Qty": 1,
                    "Subtotal": PRICING.get("GreyParrot_Monitoring_Unit", 0)
                })
            if "installation_Supervision" in PRICING and inputs.get("installation_supervision"):
                breakdown.append({
                    "Component": "Installation Supervision",
                    "Description": "Installation Supervision",
                    "Unit Price": PRICING.get("installation_Supervision", 0),
                    "Qty": 1,
                    "Subtotal": PRICING.get("installation_Supervision", 0)
                })
            if "Additional_Sorting_recipes" in PRICING and inputs.get("additional_sorting_recipes"):
                breakdown.append({
                    "Component": "Additional Sorting Recipes",
                    "Description": "Additional Sorting Recipes",
                    "Unit Price": PRICING.get("Additional_Sorting_recipes", 0),
                    "Qty": 1,
                    "Subtotal": PRICING.get("Additional_Sorting_recipes", 0)
                })
            if "SAT_to_CFA" in PRICING and inputs.get("sat_to_cfa"):
                breakdown.append({
                    "Component": "SAT to CFA",
                    "Description": "SAT to CFA",
                    "Unit Price": PRICING.get("SAT_to_CFA", 0),
                    "Qty": 1,
                    "Subtotal": PRICING.get("SAT_to_CFA", 0)
                })
            if "Engineering_&_Documentation" in PRICING and inputs.get("engineering_and_documentation"):
                breakdown.append({
                    "Component": "Engineering & Documentation",
                    "Description": "Engineering & Documentation",
                    "Unit Price": PRICING.get("Engineering_&_Documentation", 0),
                    "Qty": 1,
                    "Subtotal": PRICING.get("Engineering_&_Documentation", 0)
                })
            if "Online_Commisioning" in PRICING and inputs.get("online_commissioning"):
                breakdown.append({
                    "Component": "Online Commissioning",
                    "Description": "Online Commissioning",
                    "Unit Price": PRICING.get("Online_Commisioning", 0),
                    "Qty": 1,
                    "Subtotal": PRICING.get("Online_Commisioning", 0)
                })
            if "Installation_Commisioning_&_Training" in PRICING and inputs.get("installation_commissioning_training"):
                breakdown.append({
                    "Component": "Installation, Commissioning & Training",
                    "Description": "Installation, Commissioning & Training",
                    "Unit Price": PRICING.get("Installation_Commisioning_&_Training", 0),
                    "Qty": 1,
                    "Subtotal": PRICING.get("Installation_Commisioning_&_Training", 0)
                })
            if "LIPS2_support" in PRICING and inputs.get("lips2_support"):
                breakdown.append({
                    "Component": "LIPS2 Support",
                    "Description": "LIPS2 Support",
                    "Unit Price": PRICING.get("LIPS2_support", 0),
                    "Qty": 1,
                    "Subtotal": PRICING.get("LIPS2_support", 0)
                })

            # Key mappings for CSV
            robot_key_map = {
                "Fanuc LR-Mate": "Fanuc_LR-Mate",
                "FanucLr10iA": "FanucLr10iA",
                "Fanuc Delta DR3": "Fanuc_Delta_DR3",
                "Fanuc M10": "Fanuc_M10",
                "Fanuc M20": "Fanuc_M20",
                "Fanuc M710": "Fanuc_M710"
            }
            gripper_key_map = {
                "VentuR": "VentuR",
                "BagR": "BagR",
                "BagR CO": "BagR_CO",
                "PinchR Lr & M10": "PinchR_Lr_&_M10",
                "MonstR": "MonstR",
                "DagR": "DagR"
            }
            vision_key_map = {
                "DeepVision System": "DeepVision_System",
                "HyperVision System": "HyperVision_System"
            }

            # Robot Base key mapping (update as per your CSV keys)
            base_key_map = {
                "LrMate/Lr10ia": "LrMate/Lr10ia",
                "Delta DR3": "Delta_DR3",
                "M-10, M-20, M-710": "M10_M20_M710"
            }

            # Robot Arms (by type and quantity)
            if isinstance(inputs["robot_type"], dict):
                for rtype, qty in inputs["robot_type"].items():
                    price_key = robot_key_map.get(rtype, rtype)
                    price = PRICING.get(price_key, 0)
                    breakdown.append({
                        "Component": "Robot Arm",
                        "Description": rtype,
                        "Unit Price": price,
                        "Qty": qty,
                        "Subtotal": price * qty
                    })
            else:
                price_key = robot_key_map.get(inputs["robot_type"], inputs["robot_type"])
                price = PRICING.get(price_key, 0)
                breakdown.append({
                    "Component": "Robot Arm",
                    "Description": inputs["robot_type"],
                    "Unit Price": price,
                    "Qty": inputs["robot_arms"],
                    "Subtotal": price * inputs["robot_arms"]
                })

            # Robot Bases (by type and quantity)
            if "robot_bases" in inputs and isinstance(inputs["robot_bases"], dict):
                for btype, qty in inputs["robot_bases"].items():
                    price_key = base_key_map.get(btype, btype)
                    price = PRICING.get(price_key, 0)
                    breakdown.append({
                        "Component": "Robot Base",
                        "Description": btype,
                        "Unit Price": price,
                        "Qty": qty,
                        "Subtotal": price * qty
                    })

            # Grippers (by type and quantity)
            if isinstance(inputs["gripper_type"], dict):
                for gtype, qty in inputs["gripper_type"].items():
                    price_key = gripper_key_map.get(gtype, gtype)
                    price = PRICING.get(price_key, 0)
                    breakdown.append({
                        "Component": "Gripper",
                        "Description": gtype,
                        "Unit Price": price,
                        "Qty": qty,
                        "Subtotal": price * qty
                    })
            else:
                price_key = gripper_key_map.get(inputs["gripper_type"], inputs["gripper_type"])
                price = PRICING.get(price_key, 0)
                breakdown.append({
                    "Component": "Gripper",
                    "Description": str(inputs["gripper_type"]),
                    "Unit Price": price,
                    "Qty": 1,
                    "Subtotal": price
                })

            # Conveyor (only if present in pricing)
            if "conveyor" in PRICING and inputs["conveyor_included"] == "Yes":
                breakdown.append({
                    "Component": "Conveyor",
                    "Description": f"{inputs['conveyor_size']} inch belt",
                    "Unit Price": PRICING.get("conveyor", 0),
                    "Qty": 1,
                    "Subtotal": PRICING.get("conveyor", 0)
                })

            # Vision Systems (by type and quantity)
            if "vision_system" in inputs and isinstance(inputs["vision_system"], dict):
                for vtype, qty in inputs["vision_system"].items():
                    price_key = vision_key_map.get(vtype, vtype)
                    price = PRICING.get(price_key, 0)
                    breakdown.append({
                        "Component": "Vision System",
                        "Description": vtype,
                        "Unit Price": price,
                        "Qty": qty,
                        "Subtotal": price * qty
                    })


            if inputs.get("try_and_buy"):
                breakdown.append({
                    "Component": "Try & Buy Second Arm",
                    "Description": "Deferred Payment",
                    "Unit Price": PRICING.get("try_and_buy_arm", 0),
                    "Qty": 1,
                    "Subtotal": PRICING.get("try_and_buy_arm", 0)
                })


            # Shipping logic: by truck or by boat (container)
            shipping_method = inputs.get("shipping_method", "Truck")
            num_units = int(inputs.get("num_trucks_or_containers", 1))
            if shipping_method == "Truck":
                unit_price = PRICING.get("shipping_truck", 8250)
                desc = f"{num_units} truck(s) at ${unit_price:,.0f}/truck"
            else:
                unit_price = PRICING.get("shipping_boat_container", 11000)
                desc = f"{num_units} container(s) at ${unit_price:,.0f}/container (boat)"
            shipping_cost = unit_price * num_units
            breakdown.append({
                "Component": "Shipping",
                "Description": desc,
                "Unit Price": unit_price,
                "Qty": num_units,
                "Subtotal": shipping_cost
            })



            if inputs.get("safety_fencing"):
                breakdown.append({
                    "Component": "Safety Fencing",
                    "Description": "Robot safety fencing",
                    "Unit Price": PRICING.get("safety_fencing", 0),
                    "Qty": 1,
                    "Subtotal": PRICING.get("safety_fencing", 0)
                })

            # Warranty options
            if inputs["warranty_option"] == "1 Year (Standard)":
                breakdown.append({
                    "Component": "Warranty (1 year)",
                    "Description": "Parts + labor coverage (1 year)",
                    "Unit Price": PRICING.get("warranty_1yr", PRICING.get("warranty", 0)),
                    "Qty": 1,
                    "Subtotal": PRICING.get("warranty_1yr", PRICING.get("warranty", 0))
                })
            elif inputs["warranty_option"] == "Extended":
                breakdown.append({
                    "Component": "Warranty (Extended)",
                    "Description": "Parts + labor coverage (Extended)",
                    "Unit Price": PRICING.get("warranty_extended", 0),
                    "Qty": 1,
                    "Subtotal": PRICING.get("warranty_extended", 0)
                })

            if inputs.get("pe_stamp"):
                breakdown.append({
                    "Component": "PE Stamp",
                    "Description": "Professional engineer review",
                    "Unit Price": PRICING.get("pe_stamp", 0),
                    "Qty": 1,
                    "Subtotal": PRICING.get("pe_stamp", 0)
                })

            if inputs.get("sat"):
                breakdown.append({
                    "Component": "Site Acceptance Test (SAT)",
                    "Description": "Final performance check",
                    "Unit Price": PRICING.get("sat", 0),
                    "Qty": 1,
                    "Subtotal": PRICING.get("sat", 0)
                })

            return breakdown

        total_arms = sum(robot_type.values()) if robot_type else 0
        inputs = {
            "robot_arms": total_arms,
            "gripper_type": gripper_type,
            "try_and_buy": try_and_buy,
            "robot_type": robot_type,
            "robot_bases": robot_bases,
            "shipping_method": shipping_method,
            "num_trucks_or_containers": num_trucks_or_containers,
            "safety_fencing": safety_fencing,
            "warranty_option": warranty_option,
            "vision_system": vision_system,
            "conveyor_var_speed_license": conveyor_var_speed_license,
            "custom_ai_training": custom_ai_training,
            "robot_validator_license": robot_validator_license,
            "greyparrot_monitoring_unit": greyparrot_monitoring_unit,
            "installation_supervision": installation_supervision,
            "additional_sorting_recipes": additional_sorting_recipes,
            "sat_to_cfa": sat_to_cfa,
            "engineering_and_documentation": engineering_and_documentation,
            "online_commissioning": online_commissioning,
            "installation_commissioning_training": installation_commissioning_training,
            "lips2_support": lips2_support
        }

        df = pd.DataFrame(calculate_price_breakdown(inputs))
        multiplier = float(CURRENCY_CONVERSION.get(currency, 1.0))
        df["Unit Price"] = pd.to_numeric(df["Unit Price"], errors="coerce").fillna(0) * multiplier
        df["Subtotal"] = pd.to_numeric(df["Subtotal"], errors="coerce").fillna(0) * multiplier

        total = df["Subtotal"].sum()
        st.dataframe(df.style.format({"Unit Price": "${:,.0f}", "Subtotal": "${:,.0f}"}))
        st.markdown(f"### **Total Estimated Price: {currency} {total:,.0f}**")

        # --- DOCX Generation ---
        doc = DocxTemplate("template_practice.docx")

        # Robot Arm Images (one per selected type)
        if robot_type:
            robot_arm_images = []
            for rtype in robot_type.keys():
                base_name = rtype.lower().replace(" ", "_").replace("&", "and").replace(",", "").replace("-", "_")
                robot_arm_filename = f"robot_{base_name}.png"
                if not os.path.exists(robot_arm_filename):
                    robot_arm_filename = "robot_default.png"
                robot_arm_images.append(InlineImage(doc, robot_arm_filename, width=Mm(100), height=Mm(80)))
        else:
            robot_arm_images = [InlineImage(doc, "robot_default.png", width=Mm(100), height=Mm(80))]

        layout_image = InlineImage(doc, f"layout_{inputs['robot_arms']}_arms.png", width=Mm(100), height=Mm(80))
        layout_overview_image = InlineImage(doc, f"overview_{inputs['robot_arms']}_arms.png", width=Mm(150), height=Mm(80))
        # robot_model_image = InlineImage(doc, "fanuc_m20id25.png", width=Mm(100), height=Mm(80))

        if gripper_type:
            gripper_images = []
            for gtype in gripper_type.keys():
                base_name = gtype.lower().replace(" ", "_").replace("&", "and").replace(",", "").replace("-", "_")
                gripper_filename = f"gripper_{base_name}.png"
                if not os.path.exists(gripper_filename):
                    gripper_filename = "gripper_default.png"
                gripper_images.append(InlineImage(doc, gripper_filename, width=Mm(100), height=Mm(80)))
        else:
            gripper_images = [InlineImage(doc, "gripper_default.png", width=Mm(100), height=Mm(80))]


        # Create InlineImage for docxtpl using in-memory BytesIO
        price_table_img = InlineImage(doc, save_df_as_image(df, currency=currency), width=Mm(160))

        context = {
            "value_proposition": value_proposition,
            "application_overview": application_overview,
            "client_name": client_name,
            "client_company": client_company,
            "quote_date": quote_date.strftime("%B %d, %Y"),
            "site_location": site_location,
            "robot_type": robot_type,
            "robot_arms": inputs["robot_arms"],
            "robot_bases": sum(robot_bases.values()) if isinstance(robot_bases, dict) else robot_bases,
            "gripper_type": ", ".join(gripper_type),
            "vision_system": ", ".join(vision_system),
            "materials": ", ".join(materials),
            "belt_speed": belt_speed,
            "pick_rate": pick_rate,
            "max_object_weight": max_object_weight,
            "input_power_kva": input_power_kva,
            "avg_consumption_kw": avg_consumption_kw,
            "air_consumption_lpm": air_consumption_lpm,
            "total_price": f"{currency} {total:,.0f}",
            "warranty_option": warranty_option,
            "safety_fencing": safety_fencing,
            "try_and_buy": try_and_buy,
            "layout_image": layout_image,
            "gripper_images": gripper_images,
            "robot_arm_images": robot_arm_images,
            "layout_overview_image": layout_overview_image,
            "order_confirmation_project_kickoff": order_confirmation_project_kickoff,
            "detailed_engineering": detailed_engineering,
            "engineering_review": engineering_review,
            "procurement_fabrication": procurement_fabrication,
            "fat_shipping": fat_shipping,
            "retrofit_installation": retrofit_installation,
            "commissioning_and_SAT": commissioning_and_SAT,
            "price_table_img": price_table_img,
            #"robot_model_image": robot_model_image
        }
        

        import tempfile

        doc.render(context)

        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
            temp_file_path = tmp.name
            doc.save(temp_file_path)

        st.success("✅ Quote generated successfully!")

        with open(temp_file_path, "rb") as f:
            st.download_button(
                label="📄 Download Quote DOCX",
                data=f,
                file_name=f"{client_name}_Quote_{quote_date.strftime('%Y%m%d')}.docx",
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



