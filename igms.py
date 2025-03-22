import pandas as pd

url = "https://docs.google.com/spreadsheets/d/1AEX3jKwAdO5cROqTe6k4uNv7BCy7lPOKHrGQjZA3om0/export?format=csv&gid=<377742093>"

df = pd.read_csv(url)

print(df.head())
