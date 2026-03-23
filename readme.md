# Install dependency
pip install python-pptx

# Basic usage (title case)
python replace_allcaps_pptx.py presentation.pptx fixed.pptx

# Sentence case mode
python replace_allcaps_pptx.py presentation.pptx fixed.pptx --mode sentence

# See every change made
python replace_allcaps_pptx.py presentation.pptx fixed.pptx -v
