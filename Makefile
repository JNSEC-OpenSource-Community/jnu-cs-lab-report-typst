TYPST ?= typst
UV ?= uv
UV_CACHE_DIR ?= .cache/uv

SOURCE := report.typ
PDF := report.pdf
DOCX := report.docx
CONVERTER := scripts/typ2docx-safe.py

# 使用 typ2docx 提供的 pdf2docx 引擎导出 DOCX。
# 直接调用引擎可规避 typ2docx 0.8.0 CLI 在部分 Python 3.13 环境下的退出挂起问题。
TYP2DOCX := $(UV) run --with typ2docx --with pygments python $(CONVERTER)

.PHONY: all pdf docx clean help

all: pdf docx

pdf: $(PDF)

$(PDF): $(SOURCE)
	$(TYPST) compile $(SOURCE) $(PDF)

docx: $(DOCX)

$(DOCX): $(SOURCE) $(PDF) $(CONVERTER)
	UV_CACHE_DIR=$(UV_CACHE_DIR) $(TYP2DOCX) $(PDF) $(DOCX) $(SOURCE)

clean:
	rm -f $(PDF) $(DOCX)
	rm -rf .typ2docx .~lock.report.docx#

help:
	@echo "可用目标："
	@echo "  make pdf   - 导出 PDF"
	@echo "  make docx  - 通过 typ2docx 导出 DOCX"
	@echo "  make all   - 同时导出 PDF 和 DOCX"
	@echo "  make clean - 删除导出文件"
