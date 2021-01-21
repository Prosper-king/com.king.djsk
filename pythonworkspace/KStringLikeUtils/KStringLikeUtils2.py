# -*- coding: utf-8 -*-
import pandas as pd

import matplotlib.pyplot as plt

import seaborn as sns
from sklearn.datasets import fetch_20newsgroups
from nltk.corpus import stopwords



if __name__ == '__main__':
    dataset = fetch_20newsgroups(shuffle=True, random_state=1, remove=('header','footers','quotes'))
    documents = dataset.data
    print(dataset.target_names)
    print(len(documents))
    news_df = pd.DataFrame({'document':documents})
    # removing everything except alphabets
    news_df['clean_doc'] = news_df['document'].str.replace("[^a-zA-Z#]", " ")
    # removing short words
    news_df['clean_doc'] = news_df['clean_doc'].apply(lambda x: ' '.join([w for w in x.split() if len(w)>3]))
    stop_words = stopwords.words('english')
    # tokenization
    tokenized_doc = news_df['clean_doc'].apply(lambda x: x.split())
    # remove stop-words
    tokenized_doc = tokenized_doc.apply(lambda x: [item for item in x if item not in stop_words])