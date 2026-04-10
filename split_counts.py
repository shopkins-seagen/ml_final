import pandas as pd
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.preprocessing import LabelEncoder
from sklearn.model_selection import train_test_split, GridSearchCV
from sklearn.tree import DecisionTreeClassifier

df = pd.read_csv('data/sdtm_variables.csv').dropna()
le = LabelEncoder()
df['domain_encoded'] = le.fit_transform(df['domain'])
df['text'] = df['variable_name'].str.lower() + ' ' + df['variable_label'].str.lower()

tfidf = TfidfVectorizer(ngram_range=(1,2), max_features=200, sublinear_tf=True, stop_words='english')
X = tfidf.fit_transform(df['text']).toarray()
y = df['domain_encoded'].values
feature_names = tfidf.get_feature_names_out()

X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.25, random_state=42, stratify=y)
gs = GridSearchCV(DecisionTreeClassifier(criterion='gini', random_state=42),
                  {'max_depth':[3,5,7,10,None], 'min_samples_split':[2,4,6]},
                  cv=5, scoring='accuracy', n_jobs=-1)
gs.fit(X_train, y_train)
cart = gs.best_estimator_

tree_ = cart.tree_
split_counts = {}
for i in range(tree_.node_count):
    if tree_.feature[i] != -2:
        name = feature_names[tree_.feature[i]]
        split_counts[name] = split_counts.get(name, 0) + 1

sorted_splits = sorted(split_counts.items(), key=lambda x: -x[1])
print(f"Features used in splits : {len(sorted_splits)} out of {len(feature_names)}")
print(f"Features NOT used       : {len(feature_names) - len(sorted_splits)}")
print(f"Total split nodes       : {sum(split_counts.values())}")
print()
print(f"{'Feature':<30} {'Split Count':>12}")
print("-" * 44)
for feat, cnt in sorted_splits:
    print(f"{feat:<30} {cnt:>12}")
