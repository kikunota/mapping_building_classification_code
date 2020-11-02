import difflib



a = '山留め'
b = '山留めの設置'

seq = difflib.SequenceMatcher(None,a,b)
d = seq.ratio() * 100
print(d)
