import os.path
import base64
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google.oauth2.credentials import Credentials
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from requests import Request

import pandas as pd

# Configurações
SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']
creds = None


def save_to_csv(df=pd.DataFrame):
    df.to_csv("src/database/emails.csv")


CURR_DIR = str(os.path.dirname(os.path.realpath(__file__))+'\\')

# Carregando as credenciais
if os.path.exists(str(CURR_DIR)+'token.json'):
    creds = Credentials.from_authorized_user_file(
        str(CURR_DIR)+'token.json', SCOPES)

if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    else:
        flow = InstalledAppFlow.from_client_secrets_file(
            CURR_DIR+'credentials.json', SCOPES)
        creds = flow.run_local_server(port=0)

    with open(str(CURR_DIR)+'token.json', 'w') as token:
        token.write(creds.to_json())

# Inicializando a API do Gmail
service = build('gmail', 'v1', credentials=creds)


def preprocess_text(text):
    # Limpeza de Texto
    text = re.sub(r'<[^>]+>', '', text)  # Remove tags HTML
    text = re.sub(r'[^a-zA-Z\s]', '', text)  # Remove caracteres especiais
    text = text.strip()  # Remove espaços em branco extras

    # Tokenização
    tokens = word_tokenize(text)

    # Remoção de Stop Words
    stop_words = set(stopwords.words('portuguese'))
    tokens = [word for word in tokens if word.lower() not in stop_words]

    # Normalização de Texto (usando stemming)
    stemmer = PorterStemmer()
    tokens = [stemmer.stem(word) for word in tokens]

    # Transforme de volta em texto
    preprocessed_text = ' '.join(tokens)

    return preprocessed_text

# Função para buscar e imprimir os detalhes dos emails


def fetch_emails():
    df = pd.DataFrame(columns=['id', 'from', 'name', 'subject', 'body'])
    try:
        results = service.users().messages().list(
            userId='me', labelIds=['INBOX'], maxResults=10).execute()
        messages = results.get('messages', [])

        if not messages:
            print('Nenhum email encontrado.')
        else:
            print('Verificando os emails....:')
            c = 1
            for message in messages:
                msg = service.users().messages().get(
                    userId='me', id=message['id']).execute()
                message_data = msg['payload']['headers']
                for values in message_data:
                    name = values['name']
                    if name == 'From':
                        from_name = values['value']
                    if name == 'Subject':
                        subject = str(values['value'])
                        subject = subject if len(
                            subject) > 0 else 'Sem dados'
                #print(f'De: {from_name}, Assunto: {subject}')
                # Obtendo o corpo da mensagem
                if 'parts' in msg['payload']:
                    msg_parts = msg['payload']['parts']
                    for part in msg_parts:
                        if part['mimeType'] == 'text/plain':
                            part_data = part['body']

                            data = part_data['data']
                            byte_code = base64.urlsafe_b64decode(data)
                            text = str(byte_code.decode('utf-8'))

                            text = text if len(text) > 0 else 'Sem dados'
                            #print('Conteúdo do Email:')
                            # print(text)
                            # print("==="*40)
                df1 = pd.DataFrame(
                    [{'id': message['id'], 'from': from_name, 'name': name, 'subject': str(subject), 'body': str(text)}])

                # Adicionar Email na lista
                df = pd.concat([df, df1], axis=0, ignore_index=True)
        # Guardar os emails em CVS
        save_to_csv(df)

    except HttpError as error:
        print(f'Ocorreu um erro: { error }')
