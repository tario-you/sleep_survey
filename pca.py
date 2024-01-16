import pandas as pd
import matplotlib.pyplot as plt
from sklearn.decomposition import PCA
from sklearn.preprocessing import StandardScaler, LabelEncoder

# load excel
df = pd.read_excel("Copy of Student Sleep Survey(92).xlsx")

# drop unnecessary columns
data_for_pca = df.drop(columns=[
                       'ID', 'Start time', 'Completion time', 'Email', 'Name', 'Last modified time']).copy()

# encode categories
label_encoders = {}
for column in data_for_pca.select_dtypes(include=['object']).columns:
    le = LabelEncoder()
    data_for_pca[column] = le.fit_transform(data_for_pca[column].astype(str))
    label_encoders[column] = le

# standardize
scaler = StandardScaler()
scaled_data = scaler.fit_transform(data_for_pca)

# pca
pca = PCA(n_components=2)
principal_components = pca.fit_transform(scaled_data)

# covert pca to df
pca_df = pd.DataFrame(data=principal_components, columns=[
                      'Principal Component 1', 'Principal Component 2'])

# visualize
plt.figure(figsize=(10, 8))
plt.scatter(pca_df['Principal Component 1'],
            pca_df['Principal Component 2'], s=50)
plt.xlabel('Principal Component 1')
plt.ylabel('Principal Component 2')
plt.title('2D PCA of Student Sleep Survey Data')
plt.grid(True)
plt.show()
