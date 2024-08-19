# PPTemplate

Making a script to automate the ppt generation of image files in a repeating layout, by the request of JSquare Productions. Half of this trainwreck was courtesy of chatGPT...

## Dependencies

```python
pip install pillow pillow-heif
pip install pptx
```

## Use how

Make an images directory in the root of the project that holds the images. The folders in it are given a title slide. Any subfolders inside are given subtitle slides. Images are divided into their sets of layout ppts.

The first slide of input.pptx is title, second is subTitle and third is the image layout.

`#title` is replaced with the title folder name and `#subtitle` is replaced with the subtitle folder name.

The output is appended into the output.pptx file if it exists, or put into a new output.pptx file.

| Suggestion : keep the theme, sizes and aspect ratios of input.pptx and output.pptx the same.

## Failings

- The result for the last slide with images less than the layout format is unstable.
- Copying textboxes from input layout to output does not work properly with transferring formatting.
