# generative-ai-training

生成AI研修資料のPPTX改善・再デザイン用リポジトリ。

## PPTXサンプルの再生成

`pptx-generator/output/sample_comparison.pptx` は生成物のバイナリファイルのため、PRには含めません。PPTXが必要な場合は、次のコマンドでローカル再生成してください。

```bash
python -m pip install python-pptx
python pptx-generator/scripts/build_sample_comparison.py
```

生成コードは `pptx-generator/scripts/build_sample_comparison.py`、確認用のSVGプレビューは `pptx-generator/preview/sample_comparison.svg` で管理します。


## GitHub Actions からPPTXを生成してダウンロードする手順

`sample_comparison.pptx` はリポジトリにコミットせず、GitHub Actions の手動実行で生成して artifact からダウンロードします。

1. GitHub のリポジトリ画面で **Actions** タブを開きます。
2. 左側のワークフロー一覧から **Build sample PPTX** を選択します。
3. **Run workflow** をクリックし、対象ブランチを確認して実行します。
4. ワークフローが完了したら、実行結果ページの **Artifacts** から次の成果物をダウンロードします。
   - `sample-comparison-pptx`: `pptx-generator/output/sample_comparison.pptx`
   - `sample-comparison-preview`: `pptx-generator/preview/sample_comparison.svg`

このワークフローは `requirements.txt` の依存関係をインストールしたうえで、`pptx-generator/scripts/build_sample_comparison.py` を実行し、PPTX本体とSVGプレビューを生成します。
