## 入力したプロンプト
```plaintext
Sigmaplot による図の作成を開発して頂けませんでしょうか。

Sigmaplot 12.0 でお願いいたします。Sigmaplot 12.0 はこちらで用意致します。Windows でのみ動作します。

Sigmaplot のマニュアルはこちらにございます。
https://github.com/ywatanabe1989/SigmaPlot-v12.0-SigMacro/blob/main/SigMacro%20v1.2/references/SigmaPlotProgrammingGuide_automation.pdf

私が開発していたコードがこちらにございます。
https://github.com/ywatanabe1989/SigmaPlot-v12.0-SigMacro

具体的には
1. .vba ファイルの作成
(2. .ps1 ファイルの作成 (powershell から csv データなどの引数を渡してプロットを自動化させたいです。）)
をお願いいたします。
いずれ、Python の Matplotlib と連携させたいと思っていまして、GUI ではなく、CUI で完結させたくございます。https://github.com/ywatanabe1989/mngs/tree/main/src/mngs/plt/_subplots
```

------------------------------------------

## GPT o1 の回答

承知いたしました。Sigmaplot 12.0 を使用して、図の作成を自動化するための `.vba` ファイルおよび、PowerShell スクリプトから引数を渡してプロットを自動化するための `.ps1` ファイルの作成をお手伝いいたします。

以下に手順とサンプルコードを提供いたします。

---

### **1. `.vba` ファイルの作成**

**概要：**

- Sigmaplot の COM オートメーションを利用して、VBA スクリプトでプロットを自動化します。
- CSV ファイルからデータを読み込み、指定されたプロットを作成します。

**手順：**

1. **VBA スクリプトの作成**

   以下は、CSV ファイルを読み込んでグラフを作成する VBA スクリプトのサンプルです。

   ```vba
   ' Sigmaplot VBA Automation Script
   Sub CreateGraphFromCSV(csvPath As String, outputPath As String)
       Dim app As Object
       Set app = CreateObject("SigmaPlot.Application")
       
       ' 新しいワークシートを作成
       app.Visible = False
       app.NewWorksheet

       ' CSV ファイルをインポート
       app.CurrentWorksheet.ImportData csvPath, 1, 1, ","

       ' データをプロット
       app.NewGraph "Line Plot"
       app.CurrentGraph.SetData app.CurrentWorksheet, 1, 2  ' 列1をX、列2をYとして設定

       ' グラフを保存
       app.CurrentGraph.ExportGraph outputPath, "PNG"

       ' アプリケーションを終了
       app.Quit
       Set app = Nothing
   End Sub
   ```

   **説明：**

   - `CreateGraphFromCSV` 関数は、CSV ファイルのパスと出力する画像ファイルのパスを引数として受け取ります。
   - Sigmaplot アプリケーションを非表示で起動し、新しいワークシートを作成します。
   - CSV ファイルをインポートし、データをワークシートに読み込みます。
   - 新しいグラフを作成し、データを設定します。
   - グラフを指定したパスに画像としてエクスポートします。
   - アプリケーションを終了し、リソースを解放します。

2. **VBA スクリプトの実行方法**

   - 上記のスクリプトを `.vba` ファイルとして保存します。例：`CreateGraph.vba`
   - スクリプト内の関数を呼び出す際に、CSV ファイルのパスと出力画像のパスを指定します。

---

### **2. `.ps1` ファイルの作成 (PowerShell スクリプト)**

**概要：**

- PowerShell から VBA スクリプトを実行し、CSV データのパスや出力先などの引数を渡します。
- Sigmaplot の COM オートメーションを PowerShell から直接操作することも可能です。

**手順：**

1. **PowerShell スクリプトの作成**

   以下は、PowerShell から Sigmaplot を制御してプロットを作成するサンプルスクリプトです。

   ```powershell
   # PowerShell Script to Automate Sigmaplot
   Param(
       [string]$csvPath,
       [string]$outputPath
   )

   # Sigmaplot アプリケーションのオブジェクトを作成
   $app = New-Object -ComObject SigmaPlot.Application
   
   # アプリケーションを非表示に設定
   $app.Visible = $false
   $app.NewWorksheet()

   # CSV ファイルをインポート
   $app.CurrentWorksheet.ImportData($csvPath, 1, 1, ",")

   # データをプロット
   $app.NewGraph("Line Plot")
   $app.CurrentGraph.SetData($app.CurrentWorksheet, 1, 2)  # 列1をX、列2をYとして設定

   # グラフを保存
   $app.CurrentGraph.ExportGraph($outputPath, "PNG")

   # アプリケーションを終了
   $app.Quit()
   [System.Runtime.Interopservices.Marshal]::ReleaseComObject($app) | Out-Null
   ```

   **説明：**

   - スクリプトは、`csvPath` と `outputPath` という引数を取ります。
   - Sigmaplot の COM オブジェクトを作成し、VBA スクリプトと同様の手順でデータのインポートとプロットを行います。

2. **PowerShell スクリプトの実行方法**

   - 上記のスクリプトを `.ps1` ファイルとして保存します。例：`CreateGraph.ps1`
   - PowerShell から以下のように実行します。

     ```powershell
     .\CreateGraph.ps1 -csvPath "C:\path\to\data.csv" -outputPath "C:\path\to\output.png"
     ```

---

### **3. Python の Matplotlib との連携**

**概要：**

- 将来的に Python スクリプトから Sigmaplot の自動化を行いたい場合、`pywin32` モジュールを使用して COM オートメーションを操作できます。
- ただし、Matplotlib 自体は Python のライブラリであり、Python 内で独自にプロットを作成できます。

**手順：**

1. **Python から Sigmaplot を操作する**

   ```python
   import win32com.client

   def create_graph(csv_path, output_path):
       app = win32com.client.Dispatch("SigmaPlot.Application")
       app.Visible = False
       app.NewWorksheet()

       # CSV ファイルをインポート
       app.CurrentWorksheet.ImportData(csv_path, 1, 1, ",")

       # データをプロット
       app.NewGraph("Line Plot")
       app.CurrentGraph.SetData(app.CurrentWorksheet, 1, 2)  # 列1をX、列2をYとして設定

       # グラフを保存
       app.CurrentGraph.ExportGraph(output_path, "PNG")

       # アプリケーションを終了
       app.Quit()
       del app
   ```

   **説明：**

   - `win32com.client` を使用して、Sigmaplot の COM オブジェクトを操作します。
   - PowerShell や VBA と同様の手順で、自動化処理を行います。

2. **Matplotlib を使用してプロットを作成する**

   - Python 内で完結させたい場合、Matplotlib を使用してプロットを作成できます。
   - 既存のコードベース（GitHub のリポジトリ）を参考に、必要なプロットを作成します。

---
