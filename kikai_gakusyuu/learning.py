import glob
import pandas as pd
from sklearn.model_selection import train_test_split
import matplotlib.pyplot as plt

latest_vct = glob.glob('H:/Projects/予測/*.xlsx')[-1]
df_latest_vct = pd.read_excel(latest_vct, index_col=0)
x = df_latest_vct.iloc[:250, 1:160]
y = df_latest_vct.iloc[:250, 160:250]

x_train, x_test, y_train, y_test = train_test_split(x, y, random_state=0)
print(x_train)
print(y_train)

from sklearn.linear_model import Ridge
from sklearn import ensemble, tree

reg = ensemble.BaggingRegressor(tree.DecisionTreeRegressor(), n_estimators=100, max_samples=0.3)
y_pred_reg = reg.fit(x_train, y_train).predict(x_test)

Ridge = Ridge(alpha=0.5, random_state=0)
y_pred_Ridge = Ridge.fit(x_train, y_train).predict(x_test)


for i in range(5):
    plt.plot(y_test.iloc[i], color='red')
    plt.plot(y_pred_reg[i], color='blue')
    plt.show()
    plt.close('all')

plt.scatter(y_test, y_pred_Ridge, s=3)
plt.scatter(y_test, y_pred_reg, s=3)
plt.show()
