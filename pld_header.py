import pandas as pd
import streamlit as st
import io

# Streamlit interface for file uploads
st.title('PLD-MRID File Generator')

file1 = st.file_uploader("Upload Roaming_SC_Completion.xlsx", type=["xlsx"])
file2 = st.file_uploader("Upload Product Spec Roaming.xlsx", type=["xlsx"])

# Function to process the uploaded files and provide download link
def process_files(file1, file2):
    if file1 is not None and file2 is not None:
        # Load input files
        file1_df = pd.read_excel(file1)
        file2_df = pd.read_excel(file2)

        # Validate required columns
        required_columns_file2 = ["Keywords", "Shortcode", "Unreg", "Keyword Alias1", "Keyword Alias2", "Commercial Name", "SIM Action", "SIM Validity", "Package Validity", "Renewal", "PricePre"]
        for col in required_columns_file2:
            if col not in file2_df.columns:
                st.error(f"Missing required column '{col}' in Product Spec Roaming.xlsx")
                return

        output_file_name = None  # Initialize variable for output file name
        
        for index, row in file2_df.iterrows():
            keyword = row["Keywords"]

            # Get PO ID from file1_df based on some criteria (e.g., matching keyword)
            matching_rows = file1_df.loc[file1_df['Keyword'] == keyword, 'POID']

            if not matching_rows.empty:
                po_id_from_file1 = matching_rows.iloc[0]
                output_file_name = f"Prodef DMP-{po_id_from_file1}.xlsx"

                # Create a Pandas ExcelWriter
                with io.BytesIO() as output:
                    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
					
						# Sheet-1 "PO" sheet DataFrame
						po_df = pd.DataFrame({
							"PO ID": [po_id_from_file1],  # Matched POID from file1
							"PO Name": [po_name],  # Retrieved from file1
							"Master Keyword": [master_keyword],  # Now taken from 'Keyword' column in file1
							"Family": ["roamingSingleCountry"],  # Predefined value
							"PO Type": ["ADDON"],  # Predefined value
							"Product Category": ["b2cMobile"],  # Predefined value
							"Payment Type": ["Prepaid,Postpaid"],  # Predefined value
							"Action": ["NO_CHANGE"],  # Predefined value
						})
						po_df.to_excel(writer, sheet_name="PO", index=False)
						
						# Sheet-2 Rules-Keyword DataFrame
						keyword_master_data = {
							"PO ID": [po_id_from_file1] * 6,
							"Keyword": [
								row["Keywords"],  # 1st row
								row["Keywords"],  # 2nd row
								row["Keywords"],  # 3rd row
								"AKTIF_P26",      # 4th row
								"AKTIF",          # 5th row
								row["Unreg"]      # 6th row from file2 column "Unreg"
							],
							"Short Code": [
								str(int(row["Shortcode"])) if not pd.isna(row["Shortcode"]) else "",  # 1st row from file2 without .0
								"124",                        # 2nd row
								"929",                        # 3rd row
								"122",                        # 4th row
								"122",                        # 5th row
								"122"                         # 6th row
							],
							"Keyword Type": [
								"Master",          # 1st row
								"Master",          # 2nd row
								"Master",          # 3rd row
								"Dormant",         # 4th row
								"Dormant",         # 5th row
								"UNREG"            # 6th row
							],
							"Active": ["Yes"] * 6,
							"Routing_NGSSP": ["No"] * 6,
							"Action": ["INSERT"] * 6
						}
						keyword_master_df = pd.DataFrame(keyword_master_data)
		
						# Ensure "Short Code" exists and handle NaN values
						keyword_master_df["Short Code"] = keyword_master_df["Short Code"].astype(str).str.strip().replace("nan", "")
		
						# Save the processed DataFrame to the output Excel file
						keyword_master_df.to_excel(writer, sheet_name="Rules-Keyword", index=False)
		
						# Sheet-3 Rules-Alias DataFrame
						keyword_alias_data = {
							"PO ID": [po_id_from_file1] * 2,
							"Keyword": [
								row["Keywords"],  # 1st row
								row["Keywords"],  # 2nd row
							],
							"Short Code": [
								str(int(row["Shortcode"])),  # 1st row from file2 without .0
								str(int(row["Shortcode"])),  # 2nd row without .0
							],
							"Keyword Aliases": [
								row["Keyword Alias1"],  # 1st row
								row["Keyword Alias2"],  # 2nd row
							]
						}
						keyword_alias_df = pd.DataFrame(keyword_alias_data)
			  
						# Add the new column "Action" with the value "INSERT" for all rows
						keyword_alias_df["Action"] = "INSERT"
		
						# Save the processed DataFrame to the output Excel file
						keyword_alias_df.to_excel(writer, sheet_name="Rules-Alias", index=False)
		
						# Sheet-4 Rules-Header DataFrame
						ruleset_header_data = {
							"PO ID": [po_id_from_file1] * 3,
							"Ruleset ShortName": [""] * 3,
							"Ruleset Name": [
								row["Commercial Name"], 
								row["Commercial Name"], 
								row["Commercial Name"]
							],
							"Keyword": [row["Keywords"], "AKTIF_P26", row["Keywords"]],
							"Family": ["ROAMINGSINGLECOUNTRY"] *3,
							"Family Code": "RSC",
							"Variant Type": ["00", "00", "00"],
							"SubVariant Type": ["PRE00", "ACT00", "00000"],
							"Ruleset Version": ["1", "1", "1"],
							"Commercial Name Bahasa": [
								row["Commercial Name"], 
								row["Commercial Name"], 
								row["Commercial Name"]
							],
							"Commercial Name English": [
								row["Commercial Name"], 
								row["Commercial Name"], 
								row["Commercial Name"]
							],
							"Commercial Description": [
								row["Commercial Name"], 
								row["Commercial Name"], 
								row["Commercial Name"]
							],
							"Remarks": [""] *3,
							"Keyword Type": ["", "", ""],
							"Reference Ruleset ShortName": ["", "", ""],
							"Reference Keyword": ["", "", ""],
						}
		
						ruleset_header_df = pd.DataFrame(ruleset_header_data)
		
						# Ensure the "Short Code" column exists and manipulate it as needed
						if "Ruleset Version" in ruleset_header_df.columns:
							# Convert to string, strip whitespace, and replace "nan" with empty strings
							ruleset_header_df["Ruleset Version"] = ruleset_header_df["Ruleset Version"].astype(str).str.strip().replace("nan", "")
						else:
							# If "Ruleset Version" column is missing, create it with default empty strings
							ruleset_header_df["Ruleset Version"] = ""
		
						# Replace any NaN with empty strings explicitly
						ruleset_header_df["Ruleset Version"] = ruleset_header_df["Ruleset Version"].fillna("")
		
						# Convert to numeric, coercing errors to NaN, then fill NaN with 0 and convert to integer
						ruleset_header_df["Ruleset Version"] = pd.to_numeric(ruleset_header_df["Ruleset Version"], errors="coerce").fillna(0).astype(int)
		
						# Add the new column "Action" with the value "INSERT" for all rows
						ruleset_header_df["Action"] = "INSERT"
		
						# Save the processed DataFrame to the output Excel file
						ruleset_header_df.to_excel(writer, sheet_name="Rules-Header", index=False)
		
						# Sheet-5 Rules-PCRF
						rules_pcrf_df = pd.DataFrame(
							{
								"PO ID": ["sample"],
								"Ruleset ShortName": ["sample"],
								"SimCard Validity": ["sample"],
								"LifeTime Validity": ["sample"],
								"MaxLife Time": ["sample"],
								"UPCC Package Code": ["sample"],
								"Claim Command": ["sample"],
								"Action": ["NO_CHANGE"],
							}
						)
						rules_pcrf_df.to_excel(writer, sheet_name="PCRF", index=False)
		
						# Sheet-6 Rules-Cases-Condition
						rules_case_condition_df = pd.DataFrame(
							{
								"Ruleset ShortName": ["sample"],
								"Keyword": ["sample"],
								"OpIndex": ["sample"],
								"Rule Condition Type": ["sample"],
								"LhsOperand": ["sample"],
								"Operator": ["sample"],
								"Values": ["sample"],
								"Case Description": ["sample"],
								"Keyword Type": ["sample"],
								"Action": ["NO_CHANGE"],
							}
						)
						rules_case_condition_df.to_excel(writer, sheet_name="Rules-Cases-Condition", index=False)
		
						# Sheet-7 Rules-Cases-Success
						rules_case_success_df = pd.DataFrame(
							{
								"Ruleset ShortName": ["sample"],
								"Keyword": ["sample"],
								"OpIndex": ["sample"],
								"Effect Type": ["sample"],
								"Operator": ["sample"],
								"Values": ["sample"],
								"Exit Value": ["sample"],
								"Case Description": ["sample"],
								"Keyword Type": ["sample"],
								"Action": ["NO_CHANGE"],
							}
						)
						rules_case_success_df.to_excel(writer, sheet_name="Rules-Cases-Success", index=False)  
		
						# Sheet-8 Rules-Messages
						messages_df = pd.DataFrame(
							{
								"PO ID": ["sample"],
								"Ruleset ShortName": ["sample"],
								"Order Status": ["sample"],
								"Order Type": ["sample"],
								"Sender Address": ["sample"],
								"Channel": ["sample"],
								"Message Content Index": ["sample"],
								"Message Content": ["sample"],
								"Action": ["NO_CHANGE"],
							}
						)
						messages_df.to_excel(writer, sheet_name="Rules-Messages", index=False)
		
						# Sheet 9: Rules-Price-Mapping
						price_df = pd.DataFrame(
							{
								"Ruleset ShortName": ["sample"],
								"Variable Name": ["sample"],
								"PO ID": ["sample"],
								"Channel": ["sample"],
								"Price": ["sample"],
								"SID": ["sample"],
								"Resultant Shortname": ["sample"],
								"Action": ["NO_CHANGE"],
							}
						)
			 
						# Save the modified DataFrame to the Excel sheet
						price_df.to_excel(writer, sheet_name="Rules-Price-Mapping", index=False)
		
						# Sheet 10: Rules-Renewal
						renewal_df = pd.DataFrame(
							{
								"Ruleset ShortName": ["sample"],
								"PO ID": ["sample"],
								"Flag Auto": ["sample"],
								"Period": ["sample"],
								"Period UOM": ["sample"],
								"Flag Charge": ["sample"],
								"Flag Suspend": ["sample"],
								"Suspend Period": ["sample"],
								"Suspend UOM": ["sample"],
								"Flag Option": ["sample"],
								"Max Cycle": ["sample"],
								"Progression Renewal": ["sample"],
								"Reminder Group Id": ["sample"],
								"Amount": ["sample"],
								"Reg Subaction": ["sample"],
								"Action Failure": ["sample"],
								"Action": ["NO_CHANGE"]
							}
						)
		
						# Save the modified DataFrame to the Excel sheet
						renewal_df.to_excel(writer, sheet_name="Rules-Renewal", index=False)
		
						# Sheet 11: Rules-GSI GRP Pack
						gsi_grp_pack_df = pd.DataFrame(
							{
								"Ruleset ShortName": ["sample"],  # First row value
								"GSI GRP Pack-Group ID": ["sample"],  # First row value
								"Action": ["NO_CHANGE"],  # First row value
							}
						)
						gsi_grp_pack_df.to_excel(writer, sheet_name="Rules-GSI GRP Pack", index=False)
		
						# Sheet 12: Rules-Location Group
						location_group_df = pd.DataFrame(
							{
								"Ruleset ShortName": ["sample"],
								"Package Group": ["sample"],
								"Microcluster ID": ["sample"],
								"Action": ["NO_CHANGE"],
							}
						)
						location_group_df.to_excel(writer, sheet_name="Rules-Location Group", index=False)
		
						# Sheet 13: Rebuy-Out
						rebuy_out_df = pd.DataFrame(
							{
								"Target PO ID": ["sample"],
								"Target Ruleset ShortName": ["sample"],
								"Target MPP": ["sample"],
								"Target Group": ["sample"],
								"Service Type": ["sample"],
								"Rebuy Price": ["sample"],
								"Allow Rebuy": ["sample"],
								"Rebuy Option": ["sample"],
								"Product Family": ["sample"],
								"Source PO ID": ["sample"],
								"Source Ruleset ShortName": ["sample"],
								"Source MPP": ["sample"],
								"Source Group": ["sample"],
								"Vice Versa Consent": ["sample"],
								"Action": ["NO_CHANGE"],
							}
						)
						rebuy_out_df.to_excel(writer, sheet_name="Rebuy-Out", index=False)
		
						# Sheet 14: Rebuy-Association
						rebuy_association_df = pd.DataFrame(
							{
								"Target PO ID": ["sample"],
								"Target Ruleset ShortName": ["sample"],
								"Target MPP": ["sample"],
								"Target Group": ["sample"],
								"Service Type": ["sample"],
								"Rebuy Price": ["sample"],
								"Allow Rebuy": ["sample"],
								"Rebuy Option": ["sample"],
								"Product Family": ["sample"],
								"Source PO ID": ["sample"],
								"Source Ruleset ShortName": ["sample"],
								"Source MPP": ["sample"],
								"Source Group": ["sample"],
								"Vice Versa Consent": ["sample"],
								"Action": ["NO_CHANGE"],
							}
						)
						rebuy_association_df.to_excel(writer, sheet_name="Rebuy-Association", index=False)
		
						# Sheet 15: Incompatibility
						incompatibility_df = pd.DataFrame(
							{
								"ID": ["sample"],
								"Target PO/RulesetShortName": ["sample"],
								"Source Family": ["sample"],
								"Source PO/RulesetShortName": ["sample"],
								"Action": ["NO_CHANGE"],
							}
						)
						incompatibility_df.to_excel(writer, sheet_name="Incompatibility", index=False)
		
						# Sheet 16: Library-Addon-Name
						lib_name_df = pd.DataFrame(
							{
								"Ruleset ShortName": ["sample"],
								"PO ID": ["sample"],
								"Commercial Name": ["sample"],
								"Description": ["sample"],
								"DA": ["sample"],
								"UCUT": ["sample"],
								"Accumulator": ["sample"],
								"Master Keyword": ["sample"],
								"Master Shortcode": ["sample"],
								"Commercial Name English": ["sample"],
								"Active Period Length": ["sample"],
								"Grace Period": ["sample"],
								"Active Period Unit": ["sample"],
								"Action": ["NO_CHANGE"]
							}
						)
		
						# Save the modified DataFrame to the Excel sheet
						lib_name_df.to_excel(writer, sheet_name="Library-Addon-Name", index=False)
		
						# Sheet 17: Library-Addon-DA - copy from file3.xlsx "Library AddOn_DA"
						library_addon_da_df = pd.DataFrame(
							{
								"Ruleset ShortName": ["sample"],
								"PO ID": ["sample"],
								"Quota Name": ["sample"],
								"DA ID": ["sample"],
								"Internal Description Bahasa": ["sample"],
								"External Description Bahasa": ["sample"],
								"Internal Description English": ["sample"],
								"External Description English": ["sample"],
								"Visibility": ["sample"],
								"Custom": ["sample"],
								"Feature": ["sample"],
								"Initial Value": ["sample"],
								"Unlimited Benefit Flag": ["sample"],
								"Scenario": ["sample"],
								"Attribute Name": ["sample"],
								"Action": ["NO_CHANGE"]
							}
						)
		
						library_addon_da_df.to_excel(writer, sheet_name="Library-Addon-DA", index=False)
		
						# Sheet 18: Library-Addon-UCUT
						library_addon_ucut_df = pd.DataFrame(
							{
								"Ruleset ShortName": ["sample"],
								"PO ID": ["sample"],
								"Quota Name": ["sample"],
								"UCUT ID": ["sample"],
								"Internal Description Bahasa": ["sample"],
								"External Description Bahasa": ["sample"],
								"Internal Description English": ["sample"],
								"External Description English": ["sample"],
								"Visibility": ["sample"],
								"Custom": ["sample"],
								"Initial Value": ["sample"],
								"Unlimited Benefit Flag": ["sample"],
								"Action": ["NO_CHANGE"],
							}
						)
						library_addon_ucut_df.to_excel(writer, sheet_name="Library-Addon-UCUT", index=False)
		
						# Sheet 19: Standalone 
						standalone_df = pd.DataFrame(
							{
								"Ruleset ShortName": ["sample"],
								"PO ID": ["sample"],
								"Scenarios": ["sample"],
								"Type": ["sample"],
								"ID": ["sample"],
								"Value": ["sample"],
								"UOM": ["sample"],
								"Validity": ["sample"],
								"Provision Payload Value": ["sample"],
								"Payload Dependent Attribute": ["sample"],
								"ACTION": ["sample"],
								"Action": ["NO_CHANGE"],                }
						)
		
						standalone_df.to_excel(writer, sheet_name="Standalone", index=False)
		
						# Sheet 20: Blacklist-Gift-Promocodes
						blacklist_gift_promocodes_df = pd.DataFrame(
							[{"Ruleset ShortName": "sample", "Coherence Key": "sample", "Promo Codes": "sample", "Action": "NO_CHANGE"}]
						)
						blacklist_gift_promocodes_df.to_excel(writer, sheet_name="Blacklist-Gift-Promocodes", index=False)
		
						# Sheet 21: Blacklist-Promocodes
						blacklist_promocodes_df = pd.DataFrame(
							[{"PO ID": "sample", "Command/Keyword": "sample", "Promo Codes": "sample", "Action": "NO_CHANGE"}]
						)
						blacklist_promocodes_df.to_excel(writer, sheet_name="Blacklist-Promocodes", index=False)
		
						# Sheet 22: MYIM3-UNREG
						myim3_unreg_df = pd.DataFrame(
							[
								{
									"Ruleset ShortName": "sample",
									"Keyword": "sample",
									"Shortcode": "sample",
									"Unreg Flag": "sample",
									"Buy Extra Flag": "sample",
									"Action": "NO_CHANGE",
								}
							]
						)
						myim3_unreg_df.to_excel(writer, sheet_name="MYIM3-UNREG", index=False)
		
						# Sheet 23: ExtraPOConfig
						extrapoconfig_df = pd.DataFrame(
							[{"Ruleset ShortName": "sample", "Extra PO Keyword": "sample", "Action": "NO_CHANGE"}]
						)
						extrapoconfig_df.to_excel(writer, sheet_name="ExtraPOConfig", index=False)
		
						# Sheet 24: Keyword-Global-Variable
						keyword_global_variable_df = pd.DataFrame(
							[
								{
									"PO ID": "sample",
									"Keyword": "sample",
									"Global Variable Type": "sample",
									"Value": "sample",
									"Keyword Type": "sample",
									"Action": "NO_CHANGE",
								}
							]
						)
						keyword_global_variable_df.to_excel(writer, sheet_name="Keyword-Global-Variable", index=False)
		
						# Sheet 25: UMB-Push-Category
						umb_push_category_df = pd.DataFrame(
							[
								{
									"Ruleset ShortName": "sample",
									"Coherence Key": "sample",
									"Group Category": "sample",
									"Short Code": "sample",
									"Show Unit": "sample",
									"Action": "NO_CHANGE",
								}
							]
						)
						umb_push_category_df.to_excel(writer, sheet_name="UMB-Push-Category", index=False)
		
						# Sheet 26: Avatar-Channel
						avatar_channel_df = pd.DataFrame(
							[
								{
									"PO ID": "sample",
									"Ruleset ShortName": "sample",
									"Keyword": "sample",
									"Commercial Name": "sample",
									"Short Code": "sample",
									"PVR ID": "sample",
									"Price": "sample",
									"Action": "NO_CHANGE",
								}
							]
						)
						avatar_channel_df.to_excel(writer, sheet_name="Avatar-Channel", index=False)
		
						# Sheet 27: Dormant-Config
						dormant_config_df = pd.DataFrame(
							[{"Ruleset ShortName": "sample", "Keyword": "sample", "Short Code": "sample", "Pvr": "sample", "Action": "NO_CHANGE"}]
						)
						dormant_config_df.to_excel(writer, sheet_name="Dormant-Config", index=False)


                    # Move the file pointer to the beginning of the file so it can be downloaded
                    output.seek(0)

                    # Provide a download button for the user
                    st.download_button(
                        label=f"Download {output_file_name}",
                        data=output,
                        file_name=output_file_name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

    else:
        st.warning("Please upload both files to proceed.")

# Call the process function if both files are uploaded
if file1 is not None and file2 is not None:
    process_files(file1, file2)
