#-----------------------------------Pull Spreadsheet--------------------------------#
import time
import pandas as pd
import os
from sqlalchemy import create_engine
import numpy as np

def extract_budget_data(data_directory, archive_directory):
    os.listdir(data_directory)
    excel = np.array([data_directory + '\\' + filename for filename in os.listdir(data_directory) if filename.endswith(".xlsx")])
    for i in excel:
        LoopStartTime = time.time()
        #Store Raw Spreadsheets
        exceldata_summary = pd.read_excel(i, 'Instructions and Summary')
        exceldata_personnel = pd.read_excel(i, 'a. Personnel')
        exceldata_fringe = pd.read_excel(i, 'b. Fringe')
        exceldata_travel = pd.read_excel(i, 'c. Travel')
        exceldata_equipment = pd.read_excel(i, 'd. Equipment')
        exceldata_supplies = pd.read_excel(i, 'e. Supplies')
        exceldata_contractual = pd.read_excel(i, 'f. Contractual')
        exceldata_construction = pd.read_excel(i, 'g. Construction')
        exceldata_other = pd.read_excel(i, 'h. Other')
        exceldata_indirect = pd.read_excel(i, 'i. Indirect')
        exceldata_costshare = pd.read_excel(i, 'j. Cost Share')

        #Award Number & Awardee Name
        AwardNbr = exceldata_summary.iloc[1,1]
        Awardee = exceldata_summary.iloc[2,1]
        ModNbr = AwardNbr[-4:]
        AwardNbr = AwardNbr[:9]

        #-----------------------------------Summary Page - Budget Periods--------------------------------#
        BudgetSummaryFact = exceldata_summary.iloc[10:14,1:6]
        BudgetSummaryFact.columns = ['Budget Period', 'Federal', 'Cost Share', 'Total Costs', 'Cost Share %']
        BudgetSummaryFact['Budget Period'] = ['1', '2', '3', 'Total']
        BudgetSummaryFact = BudgetSummaryFact.reset_index(drop=True)
        BudgetSummaryFact.insert(0,'Award Number',AwardNbr)
        BudgetSummaryFact.insert(1,'Modification Number',ModNbr)
        BudgetSummaryFact = BudgetSummaryFact[BudgetSummaryFact['Total Costs'] != 0]
        BudgetSummaryFact = BudgetSummaryFact.reset_index(drop=True)
        BudgetSummaryFact = BudgetSummaryFact.drop(BudgetSummaryFact.index[BudgetSummaryFact['Total Costs'].isnull() == True])
        BudgetSummaryFact

        #-----------------------------------Summary Page - Cost Categories--------------------------------#
        BudgetCategoriesFact = exceldata_summary.iloc[16:31,0:6]
        BudgetCategoriesFact.columns = ['','Budget Period 1', 'Budget Period 2', 'Budget Period 3', 'Total Costs', '% of Project']
        BudgetCategoriesFact = BudgetCategoriesFact.T
        CategoryColumns = BudgetCategoriesFact.iloc[0,:].str.replace("a\.","").str.replace("b\.","").str.replace("c\.","").str.replace("d\.","").str.replace("e\.","").str.replace("f\.","").str.replace("g\.","").str.replace("h\.","").str.replace("i\.","")
        CategoryColumns = CategoryColumns.str.strip()
        BudgetCategoriesFact.columns = CategoryColumns
        BudgetCategoriesFact = BudgetCategoriesFact.drop(BudgetCategoriesFact.index[0])
        BudgetCategoriesFact = BudgetCategoriesFact.reset_index()
        BudgetCategoriesFact = BudgetCategoriesFact.rename(columns={'index':'Budget Period'})
        BudgetCategoriesFact['Budget Period'] = ['1', '2', '3', 'Total', '% of Project']
        BudgetCategoriesFact.insert(0,'Award Number',AwardNbr)
        BudgetCategoriesFact.insert(1,'Modification Number',ModNbr)
        BudgetCategoriesFact = BudgetCategoriesFact[BudgetCategoriesFact['Total Costs'] != 0]
        BudgetCategoriesFact = BudgetCategoriesFact.reset_index(drop=True)
        BudgetCategoriesFact = BudgetCategoriesFact.rename(columns = {'Sub-recipient':'Sub-Recipient'})
        BudgetCategoriesFact = BudgetCategoriesFact.drop(BudgetCategoriesFact.index[BudgetCategoriesFact['Budget Period'].str.find("% of Project") == 0])
        BudgetCategoriesFact = BudgetCategoriesFact.drop(BudgetCategoriesFact.index[BudgetCategoriesFact['Total Costs'].isnull() == True])
        BudgetCategoriesFact = BudgetCategoriesFact.drop(columns = ['Contractual'])
        BudgetCategoriesFact

        #---------------------------------------Personnel Page------------------------------------#

        #Import Personnel Sheet
        Personnel = exceldata_personnel.dropna(thresh=10)
        Personnel.columns = ['SOPO Task #', 'Position Title', 'Time (Hrs)', 'Pay Rate ($/Hr)', 'Total Budget Period 1', 'Time (Hrs)', 'Pay Rate ($/Hr)', 'Total Budget Period 2', 'Time (Hrs)', 'Pay Rate ($/Hr)', 'Total Budget Period 3', 'Total Project Hours', 'Project Total Dollars', 'Rate Basis']
        Personnel = Personnel.reset_index(drop=True)
        Personnel = Personnel.drop(Personnel.index[0:2])
        Personnel = Personnel.reset_index(drop=True)
        Personnel

        #Clean Personnel Data
        BudgetPeriod1 = Personnel.iloc[:,[1,0,2,3,4]]
        BudgetPeriod1.insert(0,'Budget Period',1)
        BudgetPeriod1.rename(columns={'Total Budget Period 1':'Total'}, inplace=True)
        BudgetPeriod2 = Personnel.iloc[:,[1,0,5,6,7]]
        BudgetPeriod2.insert(0,'Budget Period',2)
        BudgetPeriod2.rename(columns={'Total Budget Period 2':'Total'}, inplace=True)
        BudgetPeriod3 = Personnel.iloc[:,[1,0,8,9,10]]
        BudgetPeriod3.insert(0,'Budget Period',3)
        BudgetPeriod3.rename(columns={'Total Budget Period 3':'Total'}, inplace=True)

        PersonnelFact = BudgetPeriod1
        PersonnelFact = PersonnelFact.append(BudgetPeriod2)
        PersonnelFact = PersonnelFact.append(BudgetPeriod3)
        PersonnelFact = PersonnelFact.dropna(thresh=5)
        PersonnelFact = PersonnelFact[PersonnelFact['Total'] != 0]
        PersonnelFact.insert(0,'Award Number', AwardNbr)
        PersonnelFact.insert(1,'Modification Number',ModNbr)
        PersonnelFact = PersonnelFact.reset_index(drop=True)
        PersonnelFact

        #---------------------------------------Fringe Page------------------------------------#
        Fringe = exceldata_fringe.dropna(thresh=8).reset_index(drop=True)
        Fringe = Fringe.drop(0)
        Fringe = Fringe.drop(1)
        Fringe = Fringe.reset_index(drop=True)

        BP1Fringe = Fringe.iloc[:,[0,1,2,3]]
        BP1Fringe.columns = ['Labor Type','Personnel Costs', 'Rate', 'Total']
        BP1Fringe.insert(0,'Budget Period',1)
        BP2Fringe = Fringe.iloc[:,[0,4,5,6]]
        BP2Fringe.columns = ['Labor Type','Personnel Costs', 'Rate', 'Total']
        BP2Fringe.insert(0,'Budget Period',2)
        BP3Fringe = Fringe.iloc[:,[0,7,8,9]]
        BP3Fringe.columns = ['Labor Type','Personnel Costs', 'Rate', 'Total']
        BP3Fringe.insert(0,'Budget Period',3)

        FringeFact=pd.DataFrame(columns = ['Budget Period','Labor Type','Personnel Costs', 'Rate', 'Total'])
        FringeFact = FringeFact.append(BP1Fringe)
        FringeFact = FringeFact.append(BP2Fringe)
        FringeFact = FringeFact.append(BP3Fringe)
        FringeFact.insert(0,'Award Number', AwardNbr)
        FringeFact.insert(1,'Modification Number',ModNbr)
        FringeFact = FringeFact[FringeFact['Labor Type'] != 'Total:']
        FringeFact = FringeFact[FringeFact['Total'] != 0]
        FringeFact = FringeFact.reset_index(drop=True)
        FringeFact

        #---------------------------------------Travel Page------------------------------------#

        #BP Indices
        BP1Idx = exceldata_travel.index[exceldata_travel['Unnamed: 1'].str.find("Budget Period 1 Total") == 0]
        BP2Idx = exceldata_travel.index[exceldata_travel['Unnamed: 1'].str.find("Budget Period 2 Total") == 0]
        BP3Idx = exceldata_travel.index[exceldata_travel['Unnamed: 1'].str.find("Budget Period 3 Total") == 0]

        #BP1
        BP1Travel = exceldata_travel.iloc[0:BP1Idx[0]+1,:]
        BP1Travel = BP1Travel.reset_index(drop=True)
        BP1Travel.columns = ['SOPO Task #', 'Purpose of Travel', 'Depart From', 'Destination', 'No. of Days', 'No. of Travelers', 'Lodging per Traveler','Flight per Traveler','Vehicle per Traveler','Per Diem Per Travel', 'Cost per Trip', 'Basis for Estimating Costs']
        BP1Travel.insert(0,'Budget Period',1)

        BP1InternationalIdx = BP1Travel.index[BP1Travel['Purpose of Travel'].str.find("International Travel") == 0]
        BP1TotalIdx = BP1Travel.index[BP1Travel['Purpose of Travel'].str.find("Budget Period 1 Total") == 0]
        BP1DomesticTravel = BP1Travel.iloc[0:BP1InternationalIdx[0],:].dropna(thresh=6)
        BP1InternationalTravel = BP1Travel.iloc[BP1InternationalIdx[0]:BP1TotalIdx[0],:]

        BP1InternationalTravel = BP1InternationalTravel.dropna(thresh=4)
        BP1InternationalTravel.insert(1,'Travel Type', 'International')
        BP1InternationalTravel

        BP1DomesticTravel = BP1DomesticTravel.dropna(thresh=4)
        BP1DomesticTravel.insert(1,'Travel Type', 'Domestic')
        BP1DomesticTravel = BP1DomesticTravel.reset_index(drop=True)
        if (len(BP1DomesticTravel.index) > 0) == True:
            BP1DomesticTravel = BP1DomesticTravel.drop(0)
        BP1DomesticTravel = BP1DomesticTravel.drop(BP1DomesticTravel.index[BP1DomesticTravel['Purpose of Travel'].str.find("EXAMPLE") == 0])
        BP1DomesticTravel = BP1DomesticTravel.reset_index(drop=True)
        BP1DomesticTravel

        #BP2
        BP2Travel = exceldata_travel.iloc[BP1Idx[0]+1:BP2Idx[0]+1,:]
        BP2Travel.insert(0,'Budget Period',2)
        BP2Travel.columns = BP1Travel.columns

        BP2InternationalIdx = BP2Travel.index[BP2Travel['Purpose of Travel'].str.find("International Travel") == 0]
        BP2TotalIdx = BP2Travel.index[BP2Travel['Purpose of Travel'].str.find("Budget Period 2 Total") == 0]
        BP2DomesticTravel = BP2Travel.iloc[0:BP2InternationalIdx[0],:].dropna(thresh=5)
        BP2InternationalTravel = BP2Travel.iloc[BP2InternationalIdx[0]:BP2TotalIdx[0],:]

        BP2InternationalTravel = BP2InternationalTravel.dropna(thresh=4)
        BP2InternationalTravel.insert(1,'Travel Type', 'International')
        BP2InternationalTravel

        BP2DomesticTravel.insert(1,'Travel Type', 'Domestic')
        BP2DomesticTravel

        #BP3
        BP3Travel = exceldata_travel.iloc[BP2Idx[0]+1:BP3Idx[0]+1,:]
        BP3Travel.insert(0,'Budget Period',3)
        BP3Travel.columns = BP1Travel.columns

        BP3InternationalIdx = BP3Travel.index[BP3Travel['Purpose of Travel'].str.find("International Travel") == 0]
        BP3TotalIdx = BP3Travel.index[BP3Travel['Purpose of Travel'].str.find("Budget Period 3 Total") == 0]
        BP3DomesticTravel = BP3Travel.iloc[0:BP3InternationalIdx[0],:].dropna(thresh=6)
        BP3InternationalTravel = BP3Travel.iloc[BP3InternationalIdx[0]:BP3TotalIdx[0],:]

        BP3InternationalTravel = BP3InternationalTravel.dropna(thresh=4)
        BP3InternationalTravel.insert(1,'Travel Type', 'International')
        BP3InternationalTravel

        BP3DomesticTravel = BP3DomesticTravel.dropna(thresh=4)
        BP3DomesticTravel.insert(1,'Travel Type', 'Domestic')
        BP3DomesticTravel

        #TravelFact
        TravelFact = BP1DomesticTravel
        TravelFact = TravelFact.append(BP1InternationalTravel)
        TravelFact = TravelFact.append(BP2DomesticTravel)
        TravelFact = TravelFact.append(BP2InternationalTravel)
        TravelFact = TravelFact.append(BP3DomesticTravel)
        TravelFact = TravelFact.append(BP3InternationalTravel)
        TravelFact = TravelFact.reset_index(drop=True)
        TravelFact.insert(0,'Award Number', AwardNbr)
        TravelFact.insert(1,'Modification Number',ModNbr)
        TravelFact

        #---------------------------------------Equipment Page------------------------------------#

        #BP Indices
        BP1EquipmentIdx = exceldata_equipment.index[exceldata_equipment['Unnamed: 1'].str.find("Budget Period 1 Total") == 0]
        BP2EquipmentIdx = exceldata_equipment.index[exceldata_equipment['Unnamed: 1'].str.find("Budget Period 2 Total") == 0]
        BP3EquipmentIdx = exceldata_equipment.index[exceldata_equipment['Unnamed: 1'].str.find("Budget Period 3 Total") == 0]

        #BP1
        BP1Equipment = exceldata_equipment.iloc[0:BP1EquipmentIdx[0],:].dropna(thresh=11)
        BP1Equipment.columns = ['SOPO Task #', 'Equipment Item', 'Qty', 'Unit Cost', 'Total Cost', 'Basis of Cost', 'Justification of Need']
        BP1Equipment = BP1Equipment.reset_index(drop=True)
        BP1Equipment.insert(0,'Budget Period',1)

        #BP2
        BP2Equipment = exceldata_equipment.iloc[BP1EquipmentIdx[0]+1:BP2EquipmentIdx[0],:].dropna(thresh=11)
        BP2Equipment.columns = ['SOPO Task #', 'Equipment Item', 'Qty', 'Unit Cost', 'Total Cost', 'Basis of Cost', 'Justification of Need']
        BP2Equipment = BP2Equipment.reset_index(drop=True)
        BP2Equipment.insert(0,'Budget Period',2)

        #BP2
        BP3Equipment = exceldata_equipment.iloc[BP2EquipmentIdx[0]+1:BP3EquipmentIdx[0],:].dropna(thresh=11)
        BP3Equipment.columns = ['SOPO Task #', 'Equipment Item', 'Qty', 'Unit Cost', 'Total Cost', 'Basis of Cost', 'Justification of Need']
        BP3Equipment = BP3Equipment.reset_index(drop=True)
        BP3Equipment.insert(0,'Budget Period',3)

        EquipmentFact = BP1Equipment
        EquipmentFact = EquipmentFact.append(BP2Equipment)
        EquipmentFact = EquipmentFact.append(BP3Equipment)
        EquipmentFact = EquipmentFact.reset_index(drop=True)
        EquipmentFact = EquipmentFact.reset_index(drop=True)
        EquipmentFact.insert(0,'Award Number', AwardNbr)
        EquipmentFact.insert(1,'Modification Number',ModNbr)
        EquipmentFact

        #---------------------------------------Supplies Page------------------------------------#

        #BP Indices
        BP1SuppliesIdx = exceldata_supplies.index[exceldata_supplies['Unnamed: 1'].str.find("Budget Period 1 Total") == 0]
        BP2SuppliesIdx = exceldata_supplies.index[exceldata_supplies['Unnamed: 1'].str.find("Budget Period 2 Total") == 0]
        BP3SuppliesIdx = exceldata_supplies.index[exceldata_supplies['Unnamed: 1'].str.find("Budget Period 3 Total") == 0]

        #BP1
        BP1Supplies = exceldata_supplies.iloc[0:BP1SuppliesIdx[0],:].dropna(thresh=2)
        BP1Supplies.columns = ['SOPO Task #', 'General Category of Supplies', 'Qty', 'Unit Cost', 'Total Cost', 'Basis of Cost', 'Justification of Need']
        BP1Supplies = BP1Supplies[BP1Supplies['Total Cost'] != 0]
        BP1Supplies = BP1Supplies.reset_index(drop=True)
        BP1Supplies.insert(0,'Budget Period',1)

        #BP2
        BP2Supplies = exceldata_supplies.iloc[BP1SuppliesIdx[0]+1:BP2SuppliesIdx[0],:].dropna(thresh=2)
        BP2Supplies.columns = ['SOPO Task #', 'General Category of Supplies', 'Qty', 'Unit Cost', 'Total Cost', 'Basis of Cost', 'Justification of Need']
        BP2Supplies = BP2Supplies[BP2Supplies['Total Cost'] != 0]
        BP2Supplies = BP2Supplies.reset_index(drop=True)
        BP2Supplies.insert(0,'Budget Period',2)

        #BP3
        BP3Supplies = exceldata_supplies.iloc[BP2SuppliesIdx[0]+1:BP3SuppliesIdx[0],:].dropna(thresh=2)
        BP3Supplies.columns = ['SOPO Task #', 'General Category of Supplies', 'Qty', 'Unit Cost', 'Total Cost', 'Basis of Cost', 'Justification of Need']
        BP3Supplies = BP3Supplies[BP3Supplies['Total Cost'] != 0]
        BP3Supplies = BP3Supplies.reset_index(drop=True)
        BP3Supplies.insert(0,'Budget Period',3)

        SuppliesFact = BP1Supplies
        SuppliesFact = SuppliesFact.append(BP2Supplies)
        SuppliesFact = SuppliesFact.append(BP3Supplies)
        SuppliesFact = SuppliesFact.reset_index(drop=True)
        SuppliesFact = SuppliesFact.drop(SuppliesFact.index[SuppliesFact['SOPO Task #'].str.find("SOPO Task #") == 0])
        SuppliesFact = SuppliesFact.drop(SuppliesFact.index[SuppliesFact['General Category of Supplies'].str.find("EXAMPLE") == 0])
        SuppliesFact = SuppliesFact.reset_index(drop=True)
        SuppliesFact.insert(0,'Award Number', AwardNbr)
        SuppliesFact.insert(1,'Modification Number',ModNbr)
        SuppliesFact

        #---------------------------------------Contractual Page------------------------------------#

        #Contractual Type Indices
        SubRecipientIdx = exceldata_contractual.index[exceldata_contractual['Unnamed: 1'].str.find("Sub-Recipient") == 0]
        VendorNameIdx = exceldata_contractual.index[exceldata_contractual['Unnamed: 1'].str.find("Vendor") == 0]
        FFRDCIdx = exceldata_contractual.index[exceldata_contractual['Unnamed: 1'].str.find("FFRDC") == 0]

        #Sub-Recipient
        SubRecipient = exceldata_contractual.iloc[SubRecipientIdx[0]+1:VendorNameIdx[0],:]
        SubRecipient = SubRecipient.drop(SubRecipient.index[SubRecipient['Unnamed: 1'].str.find("EXAMPLE") == 0])
        SubRecipient = SubRecipient.drop(SubRecipient.index[SubRecipient['Unnamed: 2'].str.find("Sub-total") == 0])

        SubRecipientBP1 = SubRecipient.iloc[:,[0,1,2,3]]
        SubRecipientBP1.columns = ['SOPO Task #', 'Name/Organization', 'Purpose and Basis of Cost', 'Cost']
        SubRecipientBP2 = SubRecipient.iloc[:,[0,1,2,4]]
        SubRecipientBP2.columns = SubRecipientBP1.columns
        SubRecipientBP3 = SubRecipient.iloc[:,[0,1,2,5]]
        SubRecipientBP3.columns = SubRecipientBP1.columns
        SubRecipientBP1.insert(0,'Budget Period',1)
        SubRecipientBP2.insert(0,'Budget Period',2)
        SubRecipientBP3.insert(0,'Budget Period',3)

        SubRecipientFact = SubRecipientBP1
        SubRecipientFact = SubRecipientFact.append(SubRecipientBP2)
        SubRecipientFact = SubRecipientFact.append(SubRecipientBP3)
        SubRecipientFact = SubRecipientFact.dropna(thresh=5)
        SubRecipientFact.insert(1,'Contractor Type','Sub-Recipient')

        #Vendor
        VendorName = exceldata_contractual.iloc[VendorNameIdx[0]+1:FFRDCIdx[0],:]
        VendorName = VendorName.drop(VendorName.index[VendorName['Unnamed: 1'].str.find("EXAMPLE") == 0])
        VendorName = VendorName.drop(VendorName.index[VendorName['Unnamed: 2'].str.find("Sub-total") == 0])

        VendorBP1 = VendorName.iloc[:,[0,1,2,3]]
        VendorBP1.columns = ['SOPO Task #', 'Name/Organization', 'Purpose and Basis of Cost', 'Cost']
        VendorBP2 = VendorName.iloc[:,[0,1,2,4]]
        VendorBP2.columns = VendorBP1.columns
        VendorBP3 = VendorName.iloc[:,[0,1,2,5]]
        VendorBP3.columns = VendorBP1.columns
        VendorBP1.insert(0,'Budget Period',1)
        VendorBP2.insert(0,'Budget Period',2)
        VendorBP3.insert(0,'Budget Period',3)

        VendorFact = VendorBP1
        VendorFact = VendorFact.append(VendorBP2)
        VendorFact = VendorFact.append(VendorBP3)
        VendorFact = VendorFact.reset_index(drop=True)
        VendorFact = VendorFact.drop(VendorFact.index[VendorFact['Cost'].isnull() == True])
        VendorFact.insert(1,'Contractor Type','Vendor')
        VendorFact

        #FFRDC
        FFRDCName = exceldata_contractual.iloc[FFRDCIdx[0]+1:,:]
        FFRDCName = FFRDCName.drop(FFRDCName.index[FFRDCName['Unnamed: 1'].str.find("EXAMPLE") == 0])
        FFRDCName = FFRDCName.drop(FFRDCName.index[FFRDCName['Unnamed: 2'].str.find("Sub-total") == 0])

        FFRDCBP1 = FFRDCName.iloc[:,[0,1,2,3]]
        FFRDCBP1.columns = ['SOPO Task #', 'Name/Organization', 'Purpose and Basis of Cost', 'Cost']
        FFRDCBP2 = FFRDCName.iloc[:,[0,1,2,4]]
        FFRDCBP2.columns = FFRDCBP1.columns
        FFRDCBP3 = FFRDCName.iloc[:,[0,1,2,5]]
        FFRDCBP3.columns = FFRDCBP1.columns
        FFRDCBP1.insert(0,'Budget Period',1)
        FFRDCBP2.insert(0,'Budget Period',2)
        FFRDCBP3.insert(0,'Budget Period',3)

        FFRDCFact = FFRDCBP1
        FFRDCFact = FFRDCFact.append(FFRDCBP2)
        FFRDCFact = FFRDCFact.append(FFRDCBP3)
        FFRDCFact = FFRDCFact.dropna(thresh=5)
        FFRDCFact.insert(1,'Contractor Type','FFRDC')
        FFRDCFact

        ContractualFact = SubRecipientFact
        ContractualFact = ContractualFact.append(VendorFact)
        ContractualFact = ContractualFact.append(FFRDCFact)
        ContractualFact = ContractualFact.reset_index(drop=True)
        ContractualFact.insert(0,'Award Number', AwardNbr)
        ContractualFact.insert(1,'Modification Number',ModNbr)
        ContractualFact

        #---------------------------------------Construction Page------------------------------------#

        #BP Indices
        BP1ConstructionIdx = exceldata_construction.index[exceldata_construction['Unnamed: 1'].str.find("Budget Period 1 Total") == 0]
        BP2ConstructionIdx = exceldata_construction.index[exceldata_construction['Unnamed: 1'].str.find("Budget Period 2 Total") == 0]
        BP3ConstructionIdx = exceldata_construction.index[exceldata_construction['Unnamed: 1'].str.find("Budget Period 3 Total") == 0]

        #BP1
        BP1Construction = exceldata_construction.iloc[0:BP1ConstructionIdx[0],:].dropna(thresh=11)
        BP1Construction.columns = ['SOPO Task #', 'General Description', 'Cost', 'Basis of Cost', 'Justification of Need']
        BP1Construction = BP1Construction.reset_index(drop=True)
        BP1Construction.insert(0,'Budget Period',1)

        #BP2
        BP2Construction = exceldata_construction.iloc[BP1ConstructionIdx[0]+1:BP2ConstructionIdx[0],:].dropna(thresh=11)
        BP2Construction.columns = ['SOPO Task #', 'General Description', 'Cost', 'Basis of Cost', 'Justification of Need']
        BP2Construction = BP2Construction.reset_index(drop=True)
        BP2Construction.insert(0,'Budget Period',2)

        #BP3
        BP3Construction = exceldata_construction.iloc[BP2ConstructionIdx[0]+1:BP3ConstructionIdx[0],:].dropna(thresh=11)
        BP3Construction.columns = ['SOPO Task #', 'General Description', 'Cost', 'Basis of Cost', 'Justification of Need']
        BP3Construction = BP3Construction.reset_index(drop=True)
        BP3Construction.insert(0,'Budget Period',3)

        ConstructionFact = BP1Construction
        ConstructionFact = ConstructionFact.append(BP2Construction)
        ConstructionFact = ConstructionFact.append(BP3Construction)
        ConstructionFact = ConstructionFact.reset_index(drop=True)
        if (len(ConstructionFact.index) > 0) == True: 
            ConstructionFact = SuppliesFact.drop(0)
        ConstructionFact = ConstructionFact.reset_index(drop=True)
        ConstructionFact.insert(0,'Award Number', AwardNbr)
        ConstructionFact.insert(1,'Modification Number',ModNbr)
        ConstructionFact

        #---------------------------------------Other Page------------------------------------#

        #BP Indices
        BP1OtherIdx = exceldata_other.index[exceldata_other['Unnamed: 1'].str.find("Budget Period 1 Total") == 0]
        BP2OtherIdx = exceldata_other.index[exceldata_other['Unnamed: 1'].str.find("Budget Period 2 Total") == 0]
        BP3OtherIdx = exceldata_other.index[exceldata_other['Unnamed: 1'].str.find("Budget Period 3 Total") == 0]


        #BP1
        BP1Other = exceldata_other.iloc[0:BP1ConstructionIdx[0],:].dropna(thresh=11)
        BP1Other.columns = ['SOPO Task #', 'General Description and SOPO Task #', 'Cost', 'Basis of Cost', 'Justification of Need']
        BP1Other = BP1Other.reset_index(drop=True)
        BP1Other.insert(0,'Budget Period',1)

        #BP2
        BP2Other = exceldata_other.iloc[BP1OtherIdx[0]+1:BP3OtherIdx[0],:].dropna(thresh=11)
        BP2Other.columns = ['SOPO Task #', 'General Description and SOPO Task #', 'Cost', 'Basis of Cost', 'Justification of Need']
        BP2Other = BP2Other.reset_index(drop=True)
        BP2Other.insert(0,'Budget Period',2)

        #BP3
        BP3Other = exceldata_other.iloc[BP3OtherIdx[0]+1:BP3OtherIdx[0],:].dropna(thresh=11)
        BP3Other.columns = ['SOPO Task #', 'General Description and SOPO Task #', 'Cost', 'Basis of Cost', 'Justification of Need']
        BP3Other = BP3Other.reset_index(drop=True)
        BP3Other.insert(0,'Budget Period',3)

        OtherFact = BP1Other
        OtherFact = OtherFact.append(BP2Other)
        OtherFact = OtherFact.append(BP3Other)
        OtherFact = OtherFact.reset_index(drop=True)
        if (len(OtherFact.index) > 0) == True: 
            OtherFact = OtherFact.drop(0)
        OtherFact = OtherFact.reset_index(drop=True)
        OtherFact.insert(0,'Award Number', AwardNbr)
        OtherFact.insert(1,'Modification Number',ModNbr)
        OtherFact

        #---------------------------------------Indirect Page------------------------------------#

        exceldata_indirect = exceldata_indirect.dropna(thresh=4)

        BP1IndirectRates = exceldata_indirect.iloc[:,[0,1]]
        BP1IndirectRates.columns = ['Type','Rate/Cost']
        BP1IndirectCosts = exceldata_indirect.iloc[:,[0,1]]
        BP1IndirectCosts.columns = BP1IndirectRates.columns
        BP1IndirectFact = BP1IndirectRates.append(BP1IndirectCosts)
        BP1IndirectFact.insert(0,'Budget Period',1)

        BP2IndirectRates = exceldata_indirect.iloc[:,[0,2]]
        BP2IndirectRates.columns = ['Type','Rate/Cost']
        BP2IndirectCosts = exceldata_indirect.iloc[:,[0,2]]
        BP2IndirectCosts.columns = BP2IndirectRates.columns
        BP2IndirectFact = BP2IndirectRates.append(BP2IndirectCosts)
        BP2IndirectFact.insert(0,'Budget Period',2)

        BP3IndirectRates = exceldata_indirect.iloc[:,[0,3]]
        BP3IndirectRates.columns = ['Type','Rate/Cost']
        BP3IndirectCosts = exceldata_indirect.iloc[:,[0,3]]
        BP3IndirectCosts.columns = BP1IndirectRates.columns
        BP3IndirectFact = BP3IndirectRates.append(BP3IndirectCosts)
        BP3IndirectFact.insert(0,'Budget Period',3)

        IndirectFact = BP1IndirectFact
        IndirectFact = IndirectFact.append(BP2IndirectFact)
        IndirectFact = IndirectFact.append(BP3IndirectFact)
        IndirectFact = IndirectFact.reset_index(drop=True)
        IndirectFact = IndirectFact.drop(IndirectFact.index[IndirectFact['Rate/Cost'].str.find('Budget Period') == 0])
        IndirectFact = IndirectFact.drop(IndirectFact.index[IndirectFact['Type'].str.find('Total indirect') == 0])
        IndirectFact = IndirectFact.drop(IndirectFact.index[IndirectFact['Rate/Cost'] == 0])
        IndirectFact = IndirectFact.reset_index(drop=True)
        IndirectFact.insert(0,'Award Number', AwardNbr)
        IndirectFact.insert(1,'Modification Number',ModNbr)
        IndirectFact

        #---------------------------------------Cost Share Page------------------------------------#

        CostShare = exceldata_costshare.dropna(thresh=2)
        CostShare = CostShare.drop(CostShare.index[CostShare['Unnamed: 2'].str.find("Cost Share Item") == 0])
        CostShare = CostShare.reset_index(drop=True)
        BP1CostShare = CostShare.iloc[:,[0,1,2,3]]
        BP1CostShare.columns = ['Organization/Source', 'Type (Cash or In Kind)', 'Cost Share Item', 'Total']
        BP1CostShare.insert(0,'Budget Period', 1)
        BP2CostShare = CostShare.iloc[:,[0,1,2,4]]
        BP2CostShare.columns = ['Organization/Source', 'Type (Cash or In Kind)', 'Cost Share Item', 'Total']
        BP2CostShare.insert(0,'Budget Period', 2)
        BP3CostShare = CostShare.iloc[:,[0,1,2,5]]
        BP3CostShare.columns = ['Organization/Source', 'Type (Cash or In Kind)', 'Cost Share Item', 'Total']
        BP3CostShare.insert(0,'Budget Period', 3)

        CostShareFact = BP1CostShare
        CostShareFact = CostShareFact.append(BP2CostShare)
        CostShareFact = CostShareFact.append(BP3CostShare)
        CostShareFact = CostShareFact.reset_index(drop=True)
        CostShareFact.insert(0,'Award Number', AwardNbr)
        CostShareFact.insert(1,'Modification Number',ModNbr)
        CostShareFact = CostShareFact.drop(CostShareFact.index[CostShareFact['Cost Share Item'].str.find("Totals") == 0])
        CostShareFact = CostShareFact.drop(CostShareFact.index[CostShareFact['Organization/Source'].str.find("Total Project Cost") == 0])
        CostShareFact = CostShareFact.drop(CostShareFact.index[CostShareFact['Organization/Source'].str.find("ABC Company") == 0])
        CostShareFact = CostShareFact.drop(CostShareFact.index[CostShareFact['Total'].isnull() == True])
        CostShareFact

        #Export data to SQL
        # parameters
        DB = {'servername': 'DESKTOP-2VHPL77\SQLEXPRESS',
              'database': 'Budget',
              'driver': 'driver=SQL Server Native Client 11.0'}

        # create the connection
        engine = create_engine('mssql+pyodbc://' + DB['servername'] + '/' + DB['database'] + "?" + DB['driver'])

        #format table data types
        format_summary = {'Award Number': str ,'Modification Number': str, 'Budget Period': str, 'Federal': int, 'Cost Share': int, 'Total Costs': int, 'Cost Share %': int}
        format_categories = {'Award Number': str, 'Modification Number': str, 'Budget Period': str, 'Personnel': float, 'Fringe Benefits': float, 'Travel': float, 'Equipment': float, 'Supplies': float, 'Sub-Recipient': float, 'Vendor': float, 'FFRDC': float, 'Total Contractual': float, 'Construction': float, 'Other Direct Costs': float, 'Total Direct Costs': float, 'Indirect Charges': float, 'Total Costs': float}
        format_personnel = {'Award Number': str, 'Modification Number': str, 'Budget Period': str, 'Position Title': str, 'SOPO Task #': str, 'Time (Hrs)': float, 'Pay Rate ($/Hr)': float, 'Total': float}
        format_fringe = {'Award Number': str, 'Modification Number': str, 'Budget Period': str, 'Labor Type': str, 'Personnel Costs': float, 'Rate': float, 'Total': float}
        format_travel = {'Award Number': str, 'Modification Number': str, 'Budget Period': str, 'Travel Type': str, 'SOPO Task #': str, 'Purpose of Travel': str, 'Depart From': str, 'Destination': str, 'No. of Days': float, 'No. of Travelers': float, 'Lodging per Traveler': float, 'Flight per Traveler': float, 'Vehicle per Traveler': float, 'Per Diem Per Travel': float, 'Cost per Trip': float, 'Basis for Estimating Costs': str}
        format_equipment = {'Award Number': str, 'Modification Number': str, 'Budget Period': str, 'SOPO Task #': str, 'Equipment Item': str, 'Qty': float, 'Unit Cost': float, 'Basis of Cost': str, 'Justification of Need': str}
        format_supplies = {'Award Number': str, 'Modification Number': str, 'Budget Period': str, 'SOPO Task #': str, 'General Category of Supplies': str, 'Qty': float, 'Total Cost': float, 'Basis of Cost': str, 'Justification of Need': str}
        format_contractual = {'Award Number': str, 'Modification Number': str, 'Budget Period': str, 'Contractor Type': str, 'SOPO Task #': str, 'Name/Organization': str, 'Purpose and Basis of Cost': str,  'Cost': float}
        format_construction = {'Award Number': str, 'Modification Number': str, 'Budget Period': str, 'SOPO Task #': str, 'General Description': str, 'Cost': float, 'Basis of Cost': str, 'Justification of Need': str}
        format_other = {'Award Number': str, 'Modification Number': str, 'Budget Period': str, 'SOPO Task #': str, 'Cost': float, 'Basis of Cost': str, 'Justification of Need': str}
        format_indirect = {'Award Number': str, 'Modification Number': str, 'Budget Period': str, 'Type': str, 'Rate/Cost': float}
        format_costshare = {'Award Number': str, 'Modification Number': str, 'Budget Period': str, 'Organization/Source': str, 'Type (Cash or In Kind)': str, 'Cost Share Item': str, 'Total': float}

        #convert table formats
        BudgetSummaryFact = BudgetSummaryFact.astype(format_summary).round(2)
        BudgetCategoriesFact = BudgetCategoriesFact.astype(format_categories).round(2)
        PersonnelFact = PersonnelFact.astype(format_personnel).round(2)
        FringeFact = FringeFact.astype(format_fringe).round(2)
        TravelFact = TravelFact.astype(format_travel).round(2)
        EquipmentFact = EquipmentFact.astype(format_equipment).round(2)
        SuppliesFact = SuppliesFact.astype(format_supplies).round(2)
        ContractualFact = ContractualFact.astype(format_contractual).round(2)
        ConstructionFact = ConstructionFact.astype(format_construction).round(2)
        OtherFact = OtherFact.astype(format_other).round(2)
        IndirectFact = IndirectFact.astype(format_indirect).round(2)
        CostShareFact = CostShareFact.astype(format_costshare).round(2)

        # add fact tables to sql server
        BudgetSummaryFact.to_sql('BudgetSummaryFact', index=False, con=engine, if_exists='append')
        BudgetCategoriesFact.to_sql('BudgetCategoriesFact', index=False, con=engine, if_exists='append')
        PersonnelFact.to_sql('PersonnelFact', index=False, con=engine, if_exists='append')
        FringeFact.to_sql('FringeFact', index=False, con=engine, if_exists='append')
        TravelFact.to_sql('TravelFact', index=False, con=engine, if_exists='append')
        EquipmentFact.to_sql('EquipmentFact', index=False, con=engine, if_exists='append')
        SuppliesFact.to_sql('SuppliesFact', index=False, con=engine, if_exists='append')
        ContractualFact.to_sql('ContractualFact', index=False, con=engine, if_exists='append')
        ConstructionFact.to_sql('ConstructionFact', index=False, con=engine, if_exists='append')
        OtherFact.to_sql('OtherFact', index=False, con=engine, if_exists='append')
        IndirectFact.to_sql('IndirectFact', index=False, con=engine, if_exists='append')
        CostShareFact.to_sql('CostShareFact', index=False, con=engine, if_exists='append')

        #Move file to archive
        os.replace(i,archive+i[-25:])
        print('Completed Load of ' + i[-25:] + ' in ' + str(round(time.time()-LoopStartTime,2)) + ' seconds')
