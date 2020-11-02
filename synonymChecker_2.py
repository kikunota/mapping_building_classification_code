from gensim.models import word2vec

#モデルまでのパス
model_path = 'word2vec.gensim.model'
#モデルの読み込み
model = word2vec.Word2Vec.load(model_path)

results = model.wv.most_similar(positive=['ティーガー'])
for result in results:
    print(result)
