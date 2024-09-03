import pandas as pd
import os
import glob

# caminho para leitura dos arquivos
folder_path = os.getcwd().replace('\\src\\scripts', '\\src\\data\\raw')

# caminho do arquivo de saida
output_file = os.getcwd().replace('\\src\\scripts', '\\src\\data\\ready\\clean.xlsx')

# lista todos os arquivos de excel
excel_files = glob.glob(os.path.join(folder_path, '*.xlsx'))

if not excel_files:
  print('Nenhum arquivo compátivel encontrado')
else:
  # dataframe inicial
  dfs = []

  for excel_file in excel_files:
    try:
      # leitura do arquivo excel
      df_temp = pd.read_excel(excel_file)
      
      #obter nome do arquivo
      file_name = os.path.basename(excel_file)
      
      # criação da coluna location
      if 'brasil' in file_name.lower():
        df_temp['location'] = 'br'
      elif 'france' in file_name.lower():
        df_temp['location'] = 'fr'
      elif 'italian' in file_name.lower():
        df_temp['location'] = 'it'
        
      # criação da coluna campanha
      df_temp['campaign'] = df_temp['utm_link'].str.extract(r'utm_campaign=(.*)')
      
      # tratamento coluna data
      df_temp['sale_date'] = df_temp['sale_date'].dt.strftime('%d/%m/%Y')
      
      # criação da coluna 
      df_temp['filename'] = excel_file
      
      # remoção da coluna utm_links
      df_temp = df_temp.drop(columns=['utm_link'])
        
      dfs.append(df_temp)
    except Exception as e:
      print(f'Error ao trata o arquivo: {excel_file}, erro: {e}')
      
if dfs:
  # cocantena todas as tabelas salvas no dfs em uma unica tabela
  result = pd.concat(dfs, ignore_index=True)
  
  #configuração do motor de escrita
  writer = pd.ExcelWriter(output_file, engine='xlsxwriter')
  
  #leva os dados do resultado ao motor de excel configurado
  result.to_excel(writer, index=False, sheet_name='base')
  
  #salvar o arquivo excel
  writer._save()
  

