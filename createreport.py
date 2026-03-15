import pandas as pd
from customtkinter import filedialog
import warnings

# Suppress specific warnings (There are data validation rules in excel files, hence the warings)
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

def load_IDQ():
    global filepathIDQ
    filepathIDQ = filedialog.askopenfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
    return bool(filepathIDQ)

def load_VOL():
    global filepathVOL
    filepathVOL = filedialog.askopenfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
    return bool(filepathVOL)

def create_report(selected_marketplace, year, week):
    
    # Read data from excel file
    df_rawIDQ = pd.read_excel(filepathIDQ)
    dfIDQ = df_rawIDQ.copy()

    df_rawVOL = pd.read_excel(filepathVOL)
    dfVOL = df_rawVOL.copy()

    # Generate the dynamic path for the QA file
    qa_file_path = f'\\\\ant.amazon.com\\dept-eu\\GDN13\\RBS\\Reporting\\QA\\PL\\{year}\\W{week} quality table.xlsx'
    df_rawQA = pd.read_excel(qa_file_path)
    df_QA = df_rawQA.copy()

    # Filter data where 'IDQ_SCORE' is greater than 0
    dfIDQ = dfIDQ[dfIDQ['IDQ_SCORE'] > 0]

    # Divide by 100 to have proper percentage values
    IDQv1 = dfIDQ['IDQ_SCORE'].mean() / 100
    IDQv2 = dfIDQ['IDQ_SCORE_V2'].mean() / 100

    # Set marketplace, year and week number
    dfVOL = dfVOL[(dfVOL['MP'] == selected_marketplace) & (dfVOL['Year'] == year) & (dfVOL['Week'] == week)]

    # Determine the sum of volumes processed in certain week
    df_Key = dfVOL[dfVOL['Su-task'] == 'Keywords Backfill']
    df_ATTR = dfVOL[dfVOL['Su-task'] == 'Attributes Backfill']
    df_TCURA = dfVOL[(dfVOL['Su-task'] == 'TCU_RA') | (dfVOL['Su-task'] == 'TCU_RA Department') | (dfVOL['Su-task'] == 'TCU_RA_PC')]
    df_RAC = dfVOL[dfVOL['Su-task'] == 'RAC Backfill']
    df_TCU = dfVOL[dfVOL['Su-task'] == 'Title Clean Up']
    df_PDBP = dfVOL[dfVOL['Su-task'] == 'Attributes Backfill-Description-Bullet Point']
    df_RefBac = dfVOL[dfVOL['Su-task'] == 'Refinements Backfill']
    df_PPU = dfVOL[dfVOL['Su-task'] == 'PPU']
    df_Sherlock = dfVOL[(dfVOL['Su-task'] == 'CWS Manual Fix') | (dfVOL['Su-task'] == 'Language Detection Fix') | (dfVOL['Su-task'] == 'Language Detection PDBP Fix')]

    # Create DataFrame for all values
    values = pd.DataFrame({'Data:': ["{:.2%}".format(dfIDQ['Has_KEYWORDS'].mean()),
                                None,
                                "{:.2%}".format(IDQv2),
                                "{:.2%}".format(dfIDQ['RAC4'].mean()), 
                                "{:.2%}".format(dfIDQ['RAC10'].mean()),
                                "{:.2%}".format(IDQv1),
                                None,
                                df_Key['Total_Volumes'].sum(),
                                df_ATTR['Total_Volumes'].sum(),
                                df_TCURA['Total_Volumes'].sum(),
                                df_RAC['Total_Volumes'].sum(),
                                None, 
                                df_TCU['Total_Volumes'].sum(),
                                df_PDBP['Total_Volumes'].sum(),
                                df_RefBac['Total_Volumes'].sum(),
                                None,
                                None,
                                None,
                                None,
                                None,
                                df_PPU['Total_Volumes'].sum(),
                                df_Sherlock['Total_Volumes'].sum()
                                ]})

    # Set index for the DataFrame
    index_ = ['% Keywords coverage',
            '% QA',
            'TS IDQ v2',
            '% TS RAC 4',
            '% TS RAC 10',
            'TS IDQ (excl. Media)',
            '# ASINs Leaf node bulk',
            '# ASINs Keywords backfilled',
            '# Attributes Backfilled',
            '# TCU Attribute Backfilled',
            '# Relevat Attributes ',
            '# ASINs Categorization Backfill',
            '# ASINs Title Sanitizated',
            '# ASINs Long Product Description & Bullet Point Backfill',
            '# ASINs Refinement backfilled',
            '# WTS SIMs received',
            '# WTS SIMs processed',
            '#ASINs catalog update',
            '# ASINs suppressed',
            '# ASINs fixed & unsuppressed',
            '# PPU',
            'Sherlock']
    values.index = index_

      # Create second DataFrame
    values2 = pd.DataFrame({
        'Data:': [df_ATTR['Total_Volumes'].sum(), 
                    df_TCURA['Total_Volumes'].sum(), 
                    df_RAC['Total_Volumes'].sum(),
                    df_TCU['Total_Volumes'].sum(),
                    df_PDBP['Total_Volumes'].sum(),
                    df_Key['Total_Volumes'].sum(),
                    df_Sherlock['Total_Volumes'].sum(),
                    df_RefBac['Total_Volumes'].sum(),
                    df_PPU['Total_Volumes'].sum(),
                    df_QA['ASINs_Audited'].sum(),
                    None,
                    None,
                    None]
                    })
    
        # Set index for the second DataFrame
    index2_ = ['Attribute Backfil',
            'TCU Relevant Attributes',
            'Relevant Attributes',
            'Title Clean Up',
            'Description & Bullet Points',
            'Keywords Backfill',
            'Sherlock',
            'Refinements',
            'PPU',
            'Quality Audit Flex',
            'Manual WTS',
            'Bulk automation LNA',
            'Categorization']
    values2.index = index2_


    # Ask user where to save the file
    file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile="output", filetypes=[("Excel Files", "*.xlsx")])
    if file_path:
        # Write to Excel file with additional DataFrame underneath
        with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
            values.to_excel(writer, sheet_name='Sheet1')
            values2.to_excel(writer, sheet_name='Sheet1', startrow=len(values) + 3)  # Start after a few empty rows