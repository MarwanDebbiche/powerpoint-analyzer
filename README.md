# Powerpoint Analyzer

Tool devlopped using [dash](https://plot.ly/dash/) to analyze Powerpoint documents.


## Install dev env:

```
pip install -r requirements.txt
```

## Run app:

```
python app.py
```

**Using Docker :**

```
docker build -t ppt-analyzer .
docker run --name ppt-analyzer -p 8050:8050 --rm ppt-analyzer
```
