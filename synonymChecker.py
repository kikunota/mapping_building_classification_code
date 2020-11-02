from gensim.models import word2vec
import sys

#モデルまでのパス
model_path = 'word2vec.gensim.model'
#モデルの読み込み
model = word2vec.Word2Vec.load(model_path)

similarity = model.wv.similarity(w1="仮設", w2="足場")
print("仮設<->足場 " + str(similarity))
"""
similarity = model.wv.similarity(w1="東京", w2="千葉")
print("東京<->千葉 " + str(similarity))
similarity = model.wv.similarity(w1="東京", w2="大宮")
print("東京<->大宮 " + str(similarity))
"""
