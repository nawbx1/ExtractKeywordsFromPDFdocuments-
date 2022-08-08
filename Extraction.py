import os                        # provides functions for creating and removing a directory (folder), fetching its contents, changing and identifying the current directory
import pandas as pd              # flexible open source data analysis/manipulation tool
import glob                      # generates lists of files matching given patterns
import pdfplumber                # extracts information from .pdf documents

"""
Obtain key words from repetitive documents, then extract as a dataframe to an .xlsx !
"""

# defining the functions used in main()
def get_keyword(start, end, text):
    """
    start: should be the word prior to the keyword.
    end: should be the word that comes after the keyword.
    text: represents the text from the page(s) you've just extracted.
    """
    for i in range(len(start)):
        try:
            field = ((text.split(start[i]))[1].split(end[i])[0])
            return field
        except:
            continue

def main():
    # create an empty dataframe, from which keywords from multiple .pdf files will be later appended by rows.
    my_dataframe = pd.DataFrame()

    for files in glob.glob("/home/yeddes/Desktop/Extraction/*.pdf"):
        with pdfplumber.open(files) as pdf:
            page = pdf.pages[0]
            text = page.extract_text()
            text = " ".join(text.split())

            # use the function get_keyword as many times to get all the desired keywords from a pdf document.

            # obtain keyword #1
    
            Nom_du_fournisseur = 'BMCC'

            # obtain keyword #2
            start = ['Identifiant :']
            end = [' ']
            identifiant_unique_fourniseur = get_keyword(start, end, text)

            # obtain keyword #3
            start = ['Nom_client']
            end = ['']
            Nom_client = get_keyword(start, end, text)

            # obtain keyword #4
            start = ['Identifiant ']
            end = ['Désignation']
            identifiant_unique_du_client  = get_keyword(start, end, text)

            # obtain keyword #5
            start = ['En date du ']
            end = [' ']
            Date_facture = get_keyword(start, end, text)

            # obtain keyword #6
            start = ['Montant Total HT ']
            end = [' ']
            Montant_ht  = get_keyword(start, end, text)
            # obtain keyword #7
            start = ['TVA (13%) ']
            end = [' ']
            Montant_tva  = get_keyword(start, end, text)
            # obtain keyword #8
            start = ['Droit de Timbre ']
            end = [' ']
            Montant_du_timbre  = get_keyword(start, end, text)
             # obtain keyword #9
            start = ['Montant Total TTC ']
            end = [' ']
            Montant_TTC  = get_keyword(start, end, text)
            # create a list with the keywords extracted from current document.
            my_list = [Nom_du_fournisseur, identifiant_unique_fourniseur, Nom_client, identifiant_unique_du_client, Date_facture, Montant_ht,Montant_tva,Montant_du_timbre,Montant_TTC]

            # append my list as a row in the dataframe.
            my_list = pd.Series(my_list)

            # append the list of keywords as a row to my dataframe.
            my_dataframe = my_dataframe.append(my_list, ignore_index=True)

            print("Document's keywords have been extracted successfully!")

    # rename dataframe columns using dictionaries.
    my_dataframe = my_dataframe.rename(columns={0:'Nom_du_fournisseur',
                                                    1:'identifiant_fourniseur',
                                                    2:'Nom_client',
                                                    3:'identifiant_client',
                                                    4:'Date_facture',
                                                    5:'Montant_ht',
                                                    6:'Montant_tva',
                                                    7:'Montant_du_timbre',
                                                    8:'Montant_TTC',
                                                    })

    # change my current working directory
    save_path = ('/home/yeddes/Desktop/test')
    os.chdir(save_path)

    # extract my dataframe to an .xlsx file!
    my_dataframe.to_excel('extractionResult.xlsx', sheet_name = 'my dataframe')
    print("")
    print(my_dataframe)

if __name__ == '__main__':
    main()
