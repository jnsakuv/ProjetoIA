import pandas as pd
import spacy
from spacy.lang.pt.stop_words import STOP_WORDS
from string import punctuation
from heapq import nlargest

# Carregar o arquivo CSV
df = pd.read_csv('src/database/emails.csv')


# Função para criar o resumo de um texto
def create_summary(text):
   
    stopwords = list(STOP_WORDS)

    nlp = spacy.load('pt_core_news_sm')

    doc = nlp(text)

    tokens = [token.text for token in doc]

    custom_punctuation = punctuation + '\n'  

    word_frequencies = {}
    for word in doc:
        if word.text.lower() not in stopwords:
            if word.text.lower() not in custom_punctuation:  
                if word.text not in word_frequencies.keys():
                    word_frequencies[word.text] = 1
                else:
                    word_frequencies[word.text] += 1

    max_frequency = max(word_frequencies.values())

    for word in word_frequencies.keys():
        word_frequencies[word] = word_frequencies[word] / max_frequency

    sentence_tokens = [sent for sent in doc.sents]

    sentence_scores = {}
    for sent in sentence_tokens:
        for word in sent:
            if word.text.lower() in word_frequencies.keys():
                if sent not in sentence_scores.keys():
                    sentence_scores[sent] = word_frequencies[word.text.lower()]
                else:
                    sentence_scores[sent] += word_frequencies[word.text.lower()]

    select_length = int(len(sentence_tokens) * 0.3)

    summary = nlargest(select_length, sentence_scores, key=sentence_scores.get)

    final_summary = [word.text for word in summary]
    summary = ' '.join(final_summary)

    return summary
 # Aplicar a função create_summary à coluna "body"
df['summary'] = df['body'].apply(create_summary)

# Salvar o DataFrame com os resumos de volta em um novo arquivo CSV
df.to_csv('src/database/summary_emails.csv', index=False)
#df.to_excel('src/database/summary_emails1.xlsx', index=False)
