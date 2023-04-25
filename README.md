# MacOS&Window

## 環境安裝：

- 安裝Anaconda

```jsx
// clone anaconda env
conda activate <env_name>
conda env export > environment.yml

// 創建環境
conda env create -f environment.yml
conda activate <env_name>

//安裝環境
pip install -r requirements.txt
```

## 開啟程式：

- 到程式的路徑

```jsx
python 程式名稱
```

## 包裝變成執行檔：

```jsx
pyinstaller -F 程式名稱
```

## res
Without res folder you are not able to run the code , only Qisda member can have it

## File introduce
beta2023_3_23.py is the only code can run
test.py is used to try CSV file so that it can run successfully in beta2023_3_23.py

## UI
![image](https://github.com/leo-play-code/DDM-auto-commandline/blob/main/demo.gif)
