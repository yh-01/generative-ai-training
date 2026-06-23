# generative-ai-training

生成AI研修資料のPPTX改善・再デザイン用リポジトリ。

## PPTXサンプルの再生成

`pptx-generator/output/sample_comparison.pptx` は生成物のバイナリファイルのため、PRには含めません。PPTXが必要な場合は、次のコマンドでローカル再生成してください。

```bash
python -m pip install python-pptx
python pptx-generator/scripts/build_sample_comparison.py
```

生成コードは `pptx-generator/scripts/build_sample_comparison.py`、確認用のSVGプレビューは `pptx-generator/preview/sample_comparison.svg` で管理します。

