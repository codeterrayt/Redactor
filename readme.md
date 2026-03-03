python version used: 3.13.12
first create python environment
```
python -m venv venv
```
then activate it

```
venv\Scripts\activate
```

then install requirements
```
pip install -r requirements.txt
```

create folders: `source_ppts`, `sanitized_ppts`

put ppt files in `source_ppts` folder
then  run
```
python main.py
```
`client_mapping.json` will be created in the same directory after processing all files