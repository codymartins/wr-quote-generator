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
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
import textwrap
import re

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
GRIPPER_SPECS = {
    "VentuR": {
        "max_object_size": "300 x 300 x 100 mm",
        "min_object_size": "40 x 40 x 5 mm",
        "max_payload": "500 g"
    },
    "BagR": {
        "max_object_size": "30 l",
        "min_object_size": "1 l",
        "max_payload": "15 kg"
    },
    "PinchR": {
        "max_object_size": "300 x 300 x 300 mm",
        "min_object_size": "25 x 25 x 10 mm",
        "max_payload": "4 kg"
    },
    "MonstR": {
        "max_object_size": "500 x 500 x 300 mm",
        "min_object_size": "50 x 50 x 10 mm",
        "max_payload": "15 kg"
    }
}
ROBOT_ARM_POWER = {
    "Fanuc LRMate 200iD-7L": {"input_power_kva": 1.2, "avg_power_kw": 0.5, "air_lpm": 0},
    "FanucLr10iA": {"input_power_kva": 1.2, "avg_power_kw": 0.5, "air_lpm": 0},
    "Fanuc Delta DR3": {"input_power_kva": 12, "avg_power_kw": 2.5, "air_lpm": 0},
    "Fanuc M-10iD-10L": {"input_power_kva": 3.0, "avg_power_kw": 1, "air_lpm": 0},
    "Fanuc M-20iD-25": {"input_power_kva": 3.0, "avg_power_kw": 1, "air_lpm": 0},
    "Fanuc M-710iC-45": {"input_power_kva": 7.5, "avg_power_kw": 2.5, "air_lpm": 0},
}
GRIPPER_POWER = {
    "VentuR": {"input_power_kva": 0, "avg_power_kw": 0, "air_lpm": 12},
    "BagR": {"input_power_kva": 0, "avg_power_kw": 0, "air_lpm": 4},
    "PinchR": {"input_power_kva": 0, "avg_power_kw": 0, "air_lpm": 4},
    "MonstR": {"input_power_kva": 0, "avg_power_kw": 0, "air_lpm": 6},
}
VISION_SYSTEM_POWER = {
    "DeepVision System": {"input_power_kva": 12.975, "avg_power_kw": 5, "air_lpm": 4},
    "HyperVision System": {"input_power_kva": 12.975, "avg_power_kw": 7, "air_lpm": 15},
}

def calculate_totals(robot_type, gripper_type, vision_system):
    total_input_power_kva = 0.0
    total_avg_power_kw = 0.0
    total_air_lpm = 0

    # Robot arms
    for rtype, qty in robot_type.items():
        vals = ROBOT_ARM_POWER.get(rtype, {"input_power_kva": 0, "avg_power_kw": 0, "air_lpm": 0})
        total_input_power_kva += vals["input_power_kva"] * qty
        total_avg_power_kw += vals["avg_power_kw"] * qty
        total_air_lpm += vals["air_lpm"] * qty

    # Grippers
    for gtype, qty in gripper_type.items():
        vals = GRIPPER_POWER.get(gtype, {"input_power_kva": 0, "avg_power_kw": 0, "air_lpm": 0})
        total_input_power_kva += vals["input_power_kva"] * qty
        total_avg_power_kw += vals["avg_power_kw"] * qty
        total_air_lpm += vals["air_lpm"] * qty

    # Vision systems
    for vtype, qty in vision_system.items():
        vals = VISION_SYSTEM_POWER.get(vtype, {"input_power_kva": 0, "avg_power_kw": 0, "air_lpm": 0})
        total_input_power_kva += vals["input_power_kva"] * qty
        total_avg_power_kw += vals["avg_power_kw"] * qty
        total_air_lpm += vals["air_lpm"] * qty

    return round(total_input_power_kva, 2), round(total_avg_power_kw, 2), int(total_air_lpm)

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


# --- UI ---
tab1, tab2, tab3, tab4 = st.tabs([
    "Proposal Info", 
    "System Config", 
    "Technical Specs", 
    "Inclusions & Quote"
])

with tab1:
    st.header("Proposal Information")
    st.progress(25, text="Step 1 of 4")
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
    if currency != "CAD":
        st.markdown(
            f"Get the latest {currency}/CAD exchange rate from [xe.com](https://www.xe.com/currencyconverter/)."
        )
        user_rate = st.number_input(
            f"Enter the current {currency}/CAD exchange rate", min_value=0.0001, value=0.0001 if currency == "USD" else 0.0001, format="%.4f"
        )
        multiplier = user_rate
    else:
        multiplier = 1.0
    application_overview = st.text_area("Brief Summary of the Application")


with tab2:
    st.header("System Configuration")
    st.progress(50, text="Step 2 of 4")
    additional_arm = st.checkbox("Include Additional Arm?") 
    pick_rate = st.text_input("Pick Rate (picks/minute)")
    # Robot Arms (type and quantity)
    robot_types_list = ["Fanuc LRMate 200iD-7L", "FanucLr10iA", "Fanuc Delta DR3", "Fanuc M-10iD-10L", "Fanuc M-20iD-25", "Fanuc M-710iC-45"]
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
    gripper_types_list = ["VentuR", "BagR", "PinchR", "MonstR"] 
    selected_grippers = st.multiselect("Gripper Types", gripper_types_list)
    gripper_type = {}
    for gtype in selected_grippers:
        qty = st.number_input(f"Quantity of {gtype}", min_value=0, value=1, key=f"qty_gripper_{gtype}")
        if qty > 0:
            gripper_type[gtype] = qty

    # Backup gripper option
    add_backup_gripper = st.checkbox("Add a backup gripper?")
    if add_backup_gripper:
        backup_gripper = st.selectbox("Select backup gripper type", gripper_types_list, key="backup_gripper_type")
    else:
        backup_gripper = None

    # Consistency check
    total_arms = sum(robot_type.values()) if robot_type else 0
    total_bases = sum(robot_bases.values()) if robot_bases else 0
    total_grippers = sum(gripper_type.values()) if gripper_type else 0
    if total_arms != total_bases or total_arms != total_grippers:
        st.warning(f"⚠️ The total number of robot arms ({total_arms}), robot bases ({total_bases}), and grippers ({total_grippers}) should be the same for a valid configuration.")


with tab3:
    st.header("Technical Specs")
    st.progress(75, text="Step 3 of 4")
    vision_types_list = ["DeepVision System", "HyperVision System"]
    selected_vision_types = st.multiselect("Robot Vision System", vision_types_list)
    vision_system = {}
    for vtype in selected_vision_types:
        qty = st.number_input(f"Quantity of {vtype}", min_value=0, value=1, key=f"qty_vision_{vtype}")
        if qty > 0:
            vision_system[vtype] = qty
    # Calculate totals based on selections in Tab 2
    auto_input_power_kva, auto_avg_power_kw, auto_air_lpm = calculate_totals(robot_type, gripper_type, vision_system)

    max_object_weight = st.number_input("Maximum Object Weight per Robot (kg)", min_value=0.0)
    # Disposition prompt
    disposition = st.selectbox("Disposition", ["FTF", "IL", "N/A", "QCX"])
    # VRS Model prompt
    vrs_model = st.selectbox("VRS Model", ["900", "1200", "1600", "1800"])
    # Vision System (type and quantity)


    # Use calculated values, disable editing
    input_power_kva = st.number_input("Input Power (kVA)", min_value=0.0, value=auto_input_power_kva, disabled=True)
    avg_consumption_kw = st.number_input("Average Power Consumption (kW)", min_value=0.0, value=auto_avg_power_kw, disabled=True)
    air_consumption_lpm = st.number_input("Total Air Consumption (L/min)", min_value=0, value=auto_air_lpm, disabled=True)

with tab4:

    st.header("Inclusions & Final Quote")
    st.progress(100, text="Step 4 of 4")

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
                "Fanuc LRMate 200iD-7L": "Fanuc_LRMate_200iD-7L",
                "FanucLr10iA": "FanucLr10iA",
                "Fanuc Delta DR3": "Fanuc_Delta_DR3",
                "Fanuc M-10iD-10L": "Fanuc_M-10iD-10L",
                "Fanuc M-20iD-25": "Fanuc_M-20iD-25",
                "Fanuc M-710iC-45": "Fanuc_M-710iC-45"
            }
            gripper_key_map = {
                "VentuR": "VentuR",
                "BagR": "BagR",
                "PinchR": "PinchR",
                "MonstR": "MonstR"
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


            if inputs.get("additional_arm"):
                breakdown.append({
                    "Component": "Additional Arm",
                    "Description": "Deferred Payment",
                    "Unit Price": PRICING.get("additional_arm", 0),
                    "Qty": 1,
                    "Subtotal": PRICING.get("additional_arm", 0)
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

            # Add backup gripper if selected
            if inputs.get("add_backup_gripper") and inputs.get("backup_gripper"):
                gripper_key_map = {
                    "VentuR": "VentuR",
                    "BagR": "BagR",
                    "PinchR": "PinchR",
                    "MonstR": "MonstR"
                }
                backup_key = gripper_key_map.get(inputs["backup_gripper"], inputs["backup_gripper"])
                backup_price = PRICING.get(backup_key, 0)
                breakdown.append({
                    "Component": "Backup Gripper",
                    "Description": f"Backup: {inputs['backup_gripper']}",
                    "Unit Price": backup_price,
                    "Qty": 1,
                    "Subtotal": backup_price
                })

            return breakdown

        total_arms = sum(robot_type.values()) if robot_type else 0
        inputs = {
            "robot_arms": total_arms,
            "gripper_type": gripper_type,
            "additional_arm": additional_arm,
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
            "lips2_support": lips2_support,
            "add_backup_gripper": add_backup_gripper,
            "backup_gripper": backup_gripper,
        }

        df = pd.DataFrame(calculate_price_breakdown(inputs))
        multiplier = float(CURRENCY_CONVERSION.get(currency, 1.0))
        df["Unit Price"] = pd.to_numeric(df["Unit Price"], errors="coerce").fillna(0) * multiplier
        df["Subtotal"] = pd.to_numeric(df["Subtotal"], errors="coerce").fillna(0) * multiplier

        total = df["Subtotal"].sum()
        st.dataframe(df.style.format({"Unit Price": "${:,.0f}", "Subtotal": "${:,.0f}"}))
        st.markdown(f"### **Total Estimated Price: {currency} {total:,.0f}**")

 

        # --- Build Configuration ID for Image Lookup ---
        def sanitize(s):
            return (str(s).strip()
                    .lower()
                    .replace(" ", "_")
                    .replace("&", "and")
                    .replace(",", "")
                    .replace("-", "_")
                    .replace("/", ""))   # ✅ removes slashes

        def get_existing_config_folder(num_arms, robot_type_str, disposition_str, vrs_model_str, gripper_types_list, base_assets_path):
            """
            Try to find a config folder with the same layout but any available gripper type.
            Returns the folder path and the gripper type used.
            """
            for alt_gripper in gripper_types_list:
                alt_gripper_str = sanitize(alt_gripper)
                alt_config_id = f"{num_arms}arms_{robot_type_str}_{disposition_str}_{vrs_model_str}_{alt_gripper_str}"
                alt_folder = os.path.join(base_assets_path, alt_config_id)
                if os.path.isdir(alt_folder):
                    return alt_folder, alt_gripper
            return base_assets_path, None  # fallback to root

        num_arms = inputs["robot_arms"]

        robot_type_str = "_".join([sanitize(rt) for rt in robot_type.keys()]) if isinstance(robot_type, dict) else sanitize(robot_type)
        disposition_str = sanitize(disposition)
        vrs_model_str = sanitize(vrs_model)

        # Only use the first gripper type for config_id and images
        if isinstance(gripper_type, dict) and gripper_type:
            first_gripper = next(iter(gripper_type.keys()))
            gripper_type_str = sanitize(first_gripper)
        else:
            gripper_type_str = sanitize(gripper_type)

        config_id = f"{num_arms}arms_{robot_type_str}_{disposition_str}_{vrs_model_str}_{gripper_type_str}"

        # ✅ Force absolute path for Assets
        base_assets_path = os.path.abspath("Assets")
        assets_folder = os.path.join(base_assets_path, config_id)

        # Fallback: If the folder for the selected gripper doesn't exist, use another gripper's images for the same layout
        if not os.path.isdir(assets_folder):
            alt_folder, alt_gripper = get_existing_config_folder(
                num_arms, robot_type_str, disposition_str, vrs_model_str, gripper_types_list, base_assets_path
            )
            if alt_folder != base_assets_path:
                st.info(f"Using images from configuration with gripper '{alt_gripper}'.")
                assets_folder = alt_folder
            else:
                st.warning("No alternate configuration images found, using default images.")
                assets_folder = base_assets_path  # fallback to root

        # --- ISO Image ---
        iso_path = os.path.join(assets_folder, "iso.png")
        if not os.path.exists(iso_path):
            st.warning(f"⚠️ Missing iso.png for {config_id}, using default.")
            iso_path = "robot_default.png"

        # --- Top & Front Images ---
        top_path = os.path.join(assets_folder, "top.png")
        if not os.path.exists(top_path):
            st.warning(f"⚠️ Missing top.png for {config_id}, using default.")
            top_path = "robot_default.png"

        front_path = os.path.join(assets_folder, "front.png")
        if not os.path.exists(front_path):
            st.warning(f"⚠️ Missing front.png for {config_id}, using default.")
            front_path = "robot_default.png"



        # --- PowerPoint Generation ---
        prs = Presentation()
        prs.slide_width = Inches(8.5)
        prs.slide_height = Inches(11)
        slide_width = prs.slide_width
        slide_height = prs.slide_height

        BLUE = RGBColor(46, 125, 122)

        # Helper: Dynamically fit text to box height and width
        def fit_text_to_box(frame, text, box_height_in, box_width_in, max_font=15, min_font=8):
            from pptx.util import Pt
            # Estimate characters per line based on box width (roughly 10pt font = 12 chars/inch)
            chars_per_inch = 12
            for font_size in range(max_font, min_font - 1, -1):
                frame.clear()
                p = frame.add_paragraph()
                p.text = text
                p.font.size = Pt(font_size)
                p.font.color.rgb = RGBColor(255, 255, 255)
                p.font.name = "Arial"
                frame.word_wrap = True
                # Estimate chars per line for this font size
                cpi = chars_per_inch * (font_size / 10)
                max_line_len = int(box_width_in * cpi)
                # Estimate wrapped lines
                lines = []
                for line in text.splitlines():
                    lines.extend(textwrap.wrap(line, width=max_line_len) or [""])
                n_lines = len(lines)
                est_text_height_pt = n_lines * font_size * 1.2
                box_height_pt = box_height_in * 72
                if est_text_height_pt < box_height_pt:
                    break
        def add_page_number(slide, page_num, color=RGBColor(120, 120, 120)):
            slide.shapes.add_textbox(
                slide_width - Inches(1.2),
                slide_height - Inches(0.5),
                Inches(1),
                Inches(0.3)
            ).text_frame.text = f"{page_num}"
            tf = slide.shapes[-1].text_frame
            p = tf.paragraphs[0]
            p.font.size = Pt(12)
            p.font.color.rgb = color
            p.font.name = FONT_NAME
            p.font.italic = True
            p.alignment = 2  # Right

        def add_footer_bar(slide, text="Waste Robotics", color=BLUE):
            bar_height = Pt(4)
            bar = slide.shapes.add_shape(
                1,  # msoShapeRectangle
                0,
                slide_height - bar_height,
                slide_width,
                bar_height
            )
            bar.fill.solid()
            bar.fill.fore_color.rgb = color
            bar.line.width = Pt(0)
            bar.line.fill.background()

        def add_watermark(slide, logo_path="logo2.png"):
            if os.path.exists(logo_path):
                slide.shapes.add_picture(
                    logo_path,
                    slide_width - Inches(2.3),
                    slide_height - Inches(0.55),
                    width=Inches(2),
                    height=Inches(0.35)
                ).element.set('style', 'opacity:0.08')  # Note: python-pptx doesn't support opacity directly, but you can pre-make a transparent PNG.


        # Branding colors
        BRAND_RED = RGBColor(239, 58, 45)   # #EF3A2D
        BRAND_DARK = RGBColor(15, 15, 15)   # #0F0F0F
        WHITE = RGBColor(255, 255, 255)

        def add_branding(slide, bg_color=None):
            # Set background color (default to BRAND_DARK if not specified)
            if bg_color is None:
                bg_color = BRAND_DARK
            fill = slide.background.fill
            fill.solid()
            fill.fore_color.rgb = bg_color

            # Add logo (top left)
            slide.shapes.add_picture("logo1.png", Inches(0.2), Inches(0.2), width=Inches(1.5))

        # Always use the blank layout for new slides
        blank_layout = prs.slide_layouts[-1]  # This is usually the blank slide

        left = Inches(1)
        top = Inches(1.2)
        width = Inches(8)
        height = Inches(1.2)

        # --- Title Slide ---
        slide = prs.slides.add_slide(blank_layout)
        add_branding(slide)
        for shape in list(slide.shapes):
            if shape.is_placeholder:
                slide.shapes._spTree.remove(shape._element)

        # Place the image so its left edge is at the center of the slide
        bg_img_path = "title_background.png"
        if os.path.exists(bg_img_path):
            from PIL import Image as PILImage
            img = PILImage.open(bg_img_path)
            img_width, img_height = img.size
            slide_px_height = int(slide_height / 9525)
            # Scale image to fit slide height
            scale = slide_px_height / img_height
            new_width = int(img_width * scale)
            new_height = slide_px_height
            img_width_emu = new_width * 9525
            img_height_emu = new_height * 9525
            left = slide_width - img_width_emu  # Align right edge of image with right edge of slide
            top = 0
            slide.shapes.add_picture(
                bg_img_path,
                left,
                top,
                width=img_width_emu,
                height=img_height_emu
            )

        # Define the light blue color and font name
        LIGHT_BLUE = RGBColor(46, 125, 122)  # #2e7d7a
        FONT_NAME = "Arial"

        # Title text on left half (never overlaps image)
        left = Inches(0.25)
        top = Inches(1.2)
        width = Inches(4.25)
        height = Inches(1.2)
        title_shape = slide.shapes.add_textbox(left, top, width, height)
        title_frame = title_shape.text_frame
        title_frame.clear()
        p = title_frame.add_paragraph()
        p.text = f"Value Proposition: \n{value_proposition}"
        p.font.size = Pt(30)
        p.font.bold = False
        p.font.color.rgb = LIGHT_BLUE
        p.font.name = FONT_NAME

        # Add a thin line above the info box
        line_left = left
        line_top = top + height + Inches(0.5)
        line_width = width
        line_height = Pt(2)
        line_shape = slide.shapes.add_shape(
            1,  # msoShapeRectangle
            line_left,
            line_top,
            line_width,
            line_height
        )
        fill = line_shape.fill
        fill.solid()
        fill.fore_color.rgb = LIGHT_BLUE
        line_shape.line.color.rgb = LIGHT_BLUE
        line_shape.line.width = Pt(0)

        # Info box below title (also only left half)
        info_shape = slide.shapes.add_textbox(left, line_top + Inches(0.2), width, Inches(1.75))
        info_frame = info_shape.text_frame
        info_frame.clear()
        p = info_frame.add_paragraph()
        p.text = f"Presented to: {client_name}\nCompany: {client_company}\n\n\n\n\n\nDate: {quote_date.strftime('%B %d, %Y')}"
        p.font.size = Pt(20)
        p.font.color.rgb = LIGHT_BLUE
        p.font.name = FONT_NAME

        # --- Confidential Line and Label (bottom left) ---
        conf_left = Inches(0.25)
        conf_bottom = slide_height - Inches(1)
        conf_line_width = Inches(1.5)
        conf_line_height = Pt(2)

        # Thin red line
        conf_line_shape = slide.shapes.add_shape(
            1,  # msoShapeRectangle
            conf_left,
            conf_bottom,
            conf_line_width,
            conf_line_height
        )
        conf_line_fill = conf_line_shape.fill
        conf_line_fill.solid()
        conf_line_fill.fore_color.rgb = BRAND_RED
        conf_line_shape.line.color.rgb = BRAND_RED
        conf_line_shape.line.width = Pt(0)

        # "Confidential" label below the line
        conf_label_shape = slide.shapes.add_textbox(
            conf_left,
            conf_bottom + conf_line_height + Pt(2),
            conf_line_width,
            Inches(0.1)
        )
        conf_label_frame = conf_label_shape.text_frame
        conf_label_frame.clear()
        p = conf_label_frame.add_paragraph()
        p.text = "Confidential"
        p.font.size = Pt(14)
        p.font.bold = True
        p.font.color.rgb = BRAND_RED
        p.font.name = FONT_NAME

        # --- Application Overview Slide ---
        slide = prs.slides.add_slide(blank_layout)
        add_branding(slide, WHITE)
        for shape in list(slide.shapes):
            if shape.is_placeholder:
                slide.shapes._spTree.remove(shape._element)

        # Colors and font
        BLUE = RGBColor(46, 125, 122)  # #2e7d7a
        BRAND_DARK = RGBColor(15, 15, 15)
        FONT_NAME = "Arial"

        # Heading: Application Overview
        heading_left = Inches(0.7)
        heading_top = Inches(1.0)
        heading_width = Inches(7)
        heading_height = Inches(0.8)
        heading_shape = slide.shapes.add_textbox(heading_left, heading_top, heading_width, heading_height)
        heading_frame = heading_shape.text_frame
        heading_frame.clear()
        p = heading_frame.add_paragraph()
        p.text = "Application Overview"
        p.font.size = Pt(32)
        p.font.bold = True
        p.font.color.rgb = BLUE
        p.font.name = FONT_NAME

        # Overview text (white)
        overview_left = heading_left
        overview_top = heading_top + heading_height + Inches(0.1)
        overview_width = Inches(7)
        overview_height = Inches(1.2)
        overview_shape = slide.shapes.add_textbox(overview_left, overview_top, overview_width, overview_height)
        overview_frame = overview_shape.text_frame
        overview_frame.clear()
        overview_frame.word_wrap = True  # Ensure text wraps within the box
        p = overview_frame.add_paragraph()
        p.text = application_overview
        p.font.size = Pt(20)
        p.font.color.rgb = BRAND_DARK
        p.font.name = FONT_NAME

        # "Preliminary layout design (#arm-system)" section
        layout_label_left = heading_left
        layout_label_top = overview_top + overview_height + Inches(0.2)
        layout_label_width = Inches(7)
        layout_label_height = Inches(0.5)
        layout_label_shape = slide.shapes.add_textbox(layout_label_left, layout_label_top, layout_label_width, layout_label_height)
        layout_label_frame = layout_label_shape.text_frame
        layout_label_frame.clear()
        p = layout_label_frame.add_paragraph()
        p.text = f"Preliminary layout design ({inputs['robot_arms']}-arm system)"
        p.font.size = Pt(22)
        p.font.bold = True
        p.font.color.rgb = BLUE
        p.font.name = FONT_NAME

        # ISO image centered below the label
        iso_img_top = layout_label_top + layout_label_height + Inches(0.4)
        iso_img_width = Inches(5)
        iso_img_height = Inches(3)
        iso_img_left = heading_left 

        slide.shapes.add_picture(
            iso_path,
            iso_img_left,
            iso_img_top,
            width=iso_img_width,
            height=iso_img_height
        )
        add_page_number(slide, 2)
        add_footer_bar(slide)
        add_watermark(slide)

        # --- Layout Images Slide ---
        slide = prs.slides.add_slide(blank_layout)
        add_branding(slide, WHITE)
        for shape in list(slide.shapes):
            if shape.is_placeholder:
                slide.shapes._spTree.remove(shape._element)

        # Title: Layout Overview (#arms-system)
        layout_title = f"Layout Overview ({inputs['robot_arms']}-arm system)"
        title_left = Inches(0.7)
        title_top = Inches(1.0)
        title_width = Inches(7)
        title_height = Inches(0.8)
        title_shape = slide.shapes.add_textbox(title_left, title_top, title_width, title_height)
        title_frame = title_shape.text_frame
        title_frame.clear()
        p = title_frame.add_paragraph()
        p.text = layout_title
        p.font.size = Pt(32)
        p.font.bold = True
        p.font.color.rgb = RGBColor(46, 125, 122)  # #2e7d7a
        p.font.name = "Arial"

        def get_scaled_size(img_path, max_width_in, max_height_in):
            img = PILImage.open(img_path)
            img_w, img_h = img.size
            dpi = 96  # Assume 96 dpi for conversion
            max_w_px = max_width_in * dpi
            max_h_px = max_height_in * dpi
            scale = min(max_w_px / img_w, max_h_px / img_h, 1.0)
            return img_w * scale / dpi, img_h * scale / dpi  # return in inches

        max_img_width = 5
        max_img_height = 4
        spacing = Inches(0.5)

        # Get scaled sizes
        top_w_in, top_h_in = get_scaled_size(top_path, max_img_width, max_img_height)
        front_w_in, front_h_in = get_scaled_size(front_path, max_img_width, max_img_height)

        # Center horizontally
        img_left = Inches(1)
        img_top = title_top + title_height + Inches(0.3)

        # Top view image (top)
        slide.shapes.add_picture(
            top_path,
            img_left,
            img_top,
            width=Inches(top_w_in),
            height=Inches(top_h_in)
        )

        # Front view image (below top view)
        front_img_top = img_top + Inches(top_h_in) + spacing
        slide.shapes.add_picture(
            front_path,
            img_left,
            front_img_top,
            width=Inches(front_w_in),
            height=Inches(front_h_in)
        )

        add_page_number(slide, 3)
        add_footer_bar(slide)
        add_watermark(slide)

        # --- Robot Arm Model Slides (one per type) ---
        if robot_type:
            for idx, rtype in enumerate(robot_type.keys()):
                slide = prs.slides.add_slide(blank_layout)
                add_branding(slide, WHITE)
                for shape in list(slide.shapes):
                    if shape.is_placeholder:
                        slide.shapes._spTree.remove(shape._element)

                BLUE = RGBColor(46, 125, 122)
                FONT_NAME = "Arial"

                # Title
                title_left = Inches(0.7)
                title_top = Inches(1.0)
                title_width = Inches(7)
                title_height = Inches(0.8)
                title_shape = slide.shapes.add_textbox(title_left, title_top, title_width, title_height)
                title_frame = title_shape.text_frame
                title_frame.clear()
                p = title_frame.add_paragraph()
                p.text = f"Robot Arm Model: {rtype}"
                p.font.size = Pt(28)
                p.font.bold = True
                p.font.color.rgb = BLUE
                p.font.name = FONT_NAME

                # Image for the arm
                img_top = title_top + title_height + Inches(0.3)
                img_width = Inches(8)
                img_height = Inches(6)
                base_name = rtype.lower().replace(" ", "_").replace("&", "and").replace(",", "").replace("-", "_")
                robot_arm_filename = f"robot_{base_name}.png"
                if not os.path.exists(robot_arm_filename):
                    robot_arm_filename = "robot_default.png"
                slide.shapes.add_picture(
                    robot_arm_filename,
                    (slide_width - img_width) // 2,
                    img_top,
                    width=img_width,
                    height=img_height
                )
                add_page_number(slide, 4)
                add_footer_bar(slide)
                add_watermark(slide)

        # --- Gripper Model Slides (one per type) ---
        if gripper_type:
            for idx, gtype in enumerate(gripper_type.keys()):
                slide = prs.slides.add_slide(blank_layout)
                add_branding(slide, WHITE)
                for shape in list(slide.shapes):
                    if shape.is_placeholder:
                        slide.shapes._spTree.remove(shape._element)

                BLUE = RGBColor(46, 125, 122)
                FONT_NAME = "Arial"

                # Title
                title_left = Inches(0.7)
                title_top = Inches(1.0)
                title_width = Inches(7)
                title_height = Inches(0.8)
                title_shape = slide.shapes.add_textbox(title_left, title_top, title_width, title_height)
                title_frame = title_shape.text_frame
                title_frame.clear()
                p = title_frame.add_paragraph()
                p.text = f"Gripper Model: {gtype}"
                p.font.size = Pt(28)
                p.font.bold = True
                p.font.color.rgb = BLUE
                p.font.name = FONT_NAME

                # Image for the gripper
                img_top = title_top + title_height + Inches(0.3)
                img_width = Inches(8)
                img_height = Inches(6)
                base_name = gtype.lower().replace(" ", "_").replace("&", "and").replace(",", "").replace("-", "_")
                gripper_filename = f"gripper_{base_name}.png"
                if not os.path.exists(gripper_filename):
                    gripper_filename = "gripper_default.png"
                slide.shapes.add_picture(
                    gripper_filename,
                    (slide_width - img_width) // 2,
                    img_top,
                    width=img_width,
                    height=img_height
                )
                # Add a thin blue line beneath the image and above the text
                line_left = Inches(1.0)
                line_top = img_top + img_height + Inches(0.1)
                line_width = slide_width - 2 * line_left
                line_height = Pt(2)
                line_shape = slide.shapes.add_shape(
                    1,  # msoShapeRectangle
                    line_left,
                    line_top,
                    line_width,
                    line_height
                )
                fill = line_shape.fill
                fill.solid()
                fill.fore_color.rgb = BLUE
                line_shape.line.color.rgb = BLUE
                line_shape.line.width = Pt(0)

                # --- Add gripper info section below image ---
                info_top = line_top + Inches(0.1)
                # --- Add gripper info section below image ---
                specs = GRIPPER_SPECS.get(gtype, {})
                info_text = (
                    f"Maximum Object Size: {specs.get('max_object_size', 'N/A')}\n"
                    f"Minimum Object Size: {specs.get('min_object_size', 'N/A')}\n"
                    f"Maximum Payload: {specs.get('max_payload', 'N/A')}"
                )
                info_top = img_top + img_height + Inches(0.3)
                info_left = Inches(1.0)
                info_width = slide_width - 2 * info_left
                info_height = Inches(1.0)
                info_shape = slide.shapes.add_textbox(info_left, info_top, info_width, info_height)
                info_frame = info_shape.text_frame
                info_frame.clear()
                p = info_frame.add_paragraph()
                p.text = info_text
                p.font.size = Pt(16)
                p.font.color.rgb = BLUE
                p.font.name = FONT_NAME

                # Page number: after all arm slides
                add_page_number(slide, 5)
                add_footer_bar(slide)
                add_watermark(slide)

        # --- Vision System Sensor Fusion Slide (Simple, Conditional Images) ---
        slide = prs.slides.add_slide(blank_layout)
        add_branding(slide, WHITE)
        for shape in list(slide.shapes):
            if shape.is_placeholder:
                slide.shapes._spTree.remove(shape._element)

        BLUE = RGBColor(46, 125, 122)
        FONT_NAME = "Arial"

        # --- Centered Main Title ---
        main_title = "ROBOT VISION SYSTEM SENSOR FUSION"
        main_title_width = Inches(8)
        main_title_height = Inches(0.8)
        main_title_left = Inches(1)
        main_title_top = Inches(1.3)
        title_shape = slide.shapes.add_textbox(main_title_left, main_title_top, main_title_width, main_title_height)
        title_frame = title_shape.text_frame
        title_frame.clear()
        p = title_frame.add_paragraph()
        p.text = main_title
        p.font.size = Pt(22)
        p.font.bold = True
        p.font.color.rgb = BLUE
        p.font.name = FONT_NAME
        title_frame.paragraphs[0].alignment = 1  # Center

        # --- Insert Vision System Images Conditionally ---
        vision_imgs = []
        if "DeepVision System" in vision_system:
            vision_imgs.append(("deepvision.png", "DeepVision System"))
        if "HyperVision System" in vision_system:
            vision_imgs.append(("hypervision.png", "HyperVision System"))

        img_width = Inches(8)
        img_height = Inches(6)
        spacing = Inches(0.5)
        total_height = len(vision_imgs) * img_height + (len(vision_imgs) - 1) * spacing
        start_top = main_title_top + main_title_height + Inches(0.5)
        img_left = (slide_width - img_width) // 2

        for idx, (img_file, label) in enumerate(vision_imgs):
            img_top = start_top + idx * (img_height + spacing)
            if os.path.exists(img_file):
                slide.shapes.add_picture(
                    img_file,
                    img_left,
                    img_top,
                    width=img_width,
                    height=img_height
                )

        add_page_number(slide, 6)
        add_footer_bar(slide)
        add_watermark(slide)

        # --- Build inclusions and exclusions lists based on selections and constants ---

        # Inclusions
        inclusions_list = []
        # Robot arms
        if robot_type:
            for rtype, qty in robot_type.items():
                inclusions_list.append(f"{qty} x {rtype} robot arm(s)")
        # Robot bases
        if robot_bases:
            for btype, qty in robot_bases.items():
                inclusions_list.append(f"{qty} x robot base(s)")
        # Grippers
        if gripper_type:
            for gtype, qty in gripper_type.items():
                inclusions_list.append(f"{qty} x {gtype} gripper(s)")
        # Shipping
        inclusions_list.append(f"Shipping to {site_location}")

        # All selected inclusions from tab 5
        tab5_labels = [
            ("safety_fencing", "Safety Fencing"),
            ("conveyor_var_speed_license", "Conveyor Variable Speed License"),
            ("custom_ai_training", "Custom AI Training"),
            ("robot_validator_license", "Robot Validator License"),
            ("greyparrot_monitoring_unit", "GreyParrot Monitoring Unit"),
            ("installation_supervision", "Installation Supervision"),
            ("additional_sorting_recipes", "Additional Sorting Recipes"),
            ("sat_to_cfa", "SAT to CFA"),
            ("engineering_and_documentation", "Engineering & Documentation"),
            ("online_commissioning", "Online Commissioning"),
            ("installation_commissioning_training", "Installation, Commissioning & Training"),
            ("lips2_support", "LIPS2 Support"),
            ("warranty_option", f"Warranty: {warranty_option}" if warranty_option != "None" else None)
        ]
        for key, label in tab5_labels:
            if key == "warranty_option":
                if warranty_option != "None":
                    inclusions_list.append(label)
            elif locals().get(key):
                inclusions_list.append(label)

        inclusions_text = "\n".join(f"• {item}" for item in inclusions_list)

        # Exclusions
        exclusions_list = []
        for key, label in tab5_labels:
            if key == "warranty_option":
                if warranty_option == "None":
                    exclusions_list.append("Warranty")
            elif not locals().get(key):
                exclusions_list.append(label)

        # Always include these exclusions
        exclusions_list += [
            "All modifications required on current equipment to integrate the robotic system",
            "Electrical hookup in client’s facility",
            f"Total input power: {input_power_kva}kVA",
            f"Average Power Consumption: {avg_consumption_kw}kW",
            "Internet hookup in client’s facility (up/down 100 Mbits/sec)",
            f"Compressed air hookup in client’s facility (total air consumption: {air_consumption_lpm}L/min)",
            "Taxes, customs and/or duty charges"
        ]

        exclusions_text = "\n".join(f"• {item}" for item in exclusions_list)

        # --- Combined Specifications, Inclusions/Exclusions, and Price Slide ---
        slide = prs.slides.add_slide(blank_layout)
        add_branding(slide, WHITE)
        for shape in list(slide.shapes):
            if shape.is_placeholder:
                slide.shapes._spTree.remove(shape._element)

        BLUE = RGBColor(46, 125, 122)
        RED = RGBColor(239, 58, 45)
        BRAND_DARK = RGBColor(15, 15, 15)
        FONT_NAME = "Arial"

        # Margins and widths
        left_margin = Inches(0.7)
        content_width = slide_width - 2 * left_margin

        # --- Specifications (top) ---
        specs_top = Inches(1.2)
        specs_label_shape = slide.shapes.add_textbox(
            left_margin,
            specs_top,
            content_width,
            Inches(0.5)
        )
        specs_label_frame = specs_label_shape.text_frame
        specs_label_frame.clear()
        p = specs_label_frame.add_paragraph()
        p.text = "System Specifications"
        p.font.size = Pt(22)
        p.font.bold = True
        p.font.color.rgb = BLUE
        p.font.name = FONT_NAME

        specs_content = (
            f"Up to {pick_rate} picks/min\n"
            f"Maximum Object Weight Per Robot: {max_object_weight} kg\n"
            f"Robots operating conditions: 5°C to 45°C"
        )
        specs_content_shape = slide.shapes.add_textbox(
            left_margin,
            specs_top + Inches(0.5),
            content_width,
            Inches(0.7)
        )
        specs_content_frame = specs_content_shape.text_frame
        specs_content_frame.clear()
        p = specs_content_frame.add_paragraph()
        p.text = specs_content
        p.font.size = Pt(14)
        p.font.color.rgb = BRAND_DARK
        p.font.name = FONT_NAME

        # --- Inclusions & Exclusions (middle) ---
        mid_top = specs_top + Inches(1.65)
        mid_height = Inches(5)
        half_width = (slide_width - 2 * left_margin) / 2

        # Inclusions (left, blue background)
        inclusions_left = left_margin
        inclusions_top = mid_top
        inclusions_width = half_width - Inches(0.1)
        inclusions_height = mid_height

        left_bg = slide.shapes.add_shape(
            1,  # msoShapeRectangle
            inclusions_left,
            inclusions_top,
            inclusions_width,
            inclusions_height
        )
        fill = left_bg.fill
        fill.solid()
        fill.fore_color.rgb = BLUE
        left_bg.line.width = Pt(0)
        left_bg.line.fill.background()

        label_shape = slide.shapes.add_textbox(
            inclusions_left + Inches(0.2),
            inclusions_top,
            inclusions_width - Inches(0.4),
            Inches(0.4)
        )
        label_frame = label_shape.text_frame
        label_frame.clear()
        p = label_frame.add_paragraph()
        p.text = "Inclusions"
        p.font.size = Pt(18)
        p.font.bold = True
        p.font.color.rgb = WHITE
        p.font.name = FONT_NAME

        inclusions_box = slide.shapes.add_textbox(
            inclusions_left + Inches(0.2),
            inclusions_top + Inches(0.5),
            inclusions_width - Inches(0.2),
            inclusions_height - Inches(0.3)
        )
        inclusions_frame = inclusions_box.text_frame
        inclusions_frame.clear()
        p = inclusions_frame.add_paragraph()
        p.text = inclusions_text
        p.font.size = Pt(12)
        p.font.color.rgb = WHITE
        p.font.name = FONT_NAME

        # Exclusions (right, red background)
        exclusions_left = left_margin + half_width + Inches(0.1)
        exclusions_top = mid_top
        exclusions_width = half_width - Inches(0.1)
        exclusions_height = mid_height

        right_bg = slide.shapes.add_shape(
            1,  # msoShapeRectangle
            exclusions_left,
            exclusions_top,
            exclusions_width,
            exclusions_height
        )
        fill = right_bg.fill
        fill.solid()
        fill.fore_color.rgb = RED
        right_bg.line.width = Pt(0)
        right_bg.line.fill.background()

        ex_label_shape = slide.shapes.add_textbox(
            exclusions_left + Inches(0.2),
            exclusions_top,
            exclusions_width - Inches(0.4),
            Inches(0.4)
        )
        ex_label_frame = ex_label_shape.text_frame
        ex_label_frame.clear()
        p = ex_label_frame.add_paragraph()
        p.text = "Exclusions"
        p.font.size = Pt(18)
        p.font.bold = True
        p.font.color.rgb = WHITE
        p.font.name = FONT_NAME

        exclusions_box = slide.shapes.add_textbox(
            exclusions_left + Inches(0.2),
            exclusions_top + Inches(0.5),
            exclusions_width - Inches(0.2),
            exclusions_height - Inches(0.3)
        )
        exclusions_frame = exclusions_box.text_frame
        exclusions_frame.clear()
        exclusions_frame.word_wrap = True
        for item in exclusions_list:
            p = exclusions_frame.add_paragraph()
            p.text = f"• {item}"
            p.font.size = Pt(12)
            p.font.color.rgb = WHITE
            p.font.name = FONT_NAME

        # --- Price Info (bottom) ---
        price_top = mid_top + mid_height - Inches(0.25)
        price_label_shape = slide.shapes.add_textbox(
            left_margin,
            price_top,
            content_width,
            Inches(0.5)
        )
        price_label_frame = price_label_shape.text_frame
        price_label_frame.clear()
        p = price_label_frame.add_paragraph()
        p.text = "Buying Price"
        p.font.size = Pt(22)
        p.font.bold = True
        p.font.color.rgb = BLUE
        p.font.name = FONT_NAME

        additional_arm_price = PRICING.get("additional_arm", 0) * multiplier
        price_content = (
            f"Robotic Sorting System: {currency} {total:,.0f}\n"
            f"Additional Robot Arm: {currency} {additional_arm_price:,.0f}"
        )
        price_content_shape = slide.shapes.add_textbox(
            left_margin,
            price_top + Inches(0.6),
            content_width,
            Inches(0.7)
        )
        price_content_frame = price_content_shape.text_frame
        price_content_frame.clear()
        p = price_content_frame.add_paragraph()
        p.text = price_content
        p.font.size = Pt(14)
        p.font.color.rgb = BRAND_DARK
        p.font.name = FONT_NAME

        # Disclaimer in small font
        disclaimer = "* Prices may vary due to exchange rates, inflation, and integration engineering. Valid for 30 days."
        disclaimer_shape = slide.shapes.add_textbox(
            left_margin,
            price_top + Inches(1.1),
            content_width,
            Inches(0.3)
        )
        disclaimer_frame = disclaimer_shape.text_frame
        disclaimer_frame.clear()
        p = disclaimer_frame.add_paragraph()
        p.text = disclaimer
        p.font.size = Pt(10)
        p.font.color.rgb = BRAND_DARK
        p.font.name = FONT_NAME

        add_page_number(slide, 7)
        add_footer_bar(slide)
        add_watermark(slide)

        delivery_top = price_top + Inches(1.5)
        deliver_shape = slide.shapes.add_textbox(
            left_margin,
            delivery_top,
            content_width,
            Inches(0.5)
        )
        delivery_frame = deliver_shape.text_frame
        delivery_frame.clear()
        p = delivery_frame.add_paragraph()
        p.text = "Estimated Delivery Time: 24 Weeks (Subject to Change)"
        p.font.size = Pt(16)
        p.font.bold = True
        p.font.color.rgb = BRAND_DARK
        p.font.name = FONT_NAME

        # --- Download PPTX ---
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pptx") as tmp:
            pptx_file_path = tmp.name
            prs.save(pptx_file_path)

        st.success("✅ Quote PowerPoint generated successfully!")

        with open(pptx_file_path, "rb") as f:
            st.download_button(
                label="📊 Download Quote PPTX",
                data=f,
                file_name=f"{client_name}_Quote_{quote_date.strftime('%Y%m%d')}.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )


# Footer branding (if needed)
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



