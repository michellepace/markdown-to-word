# Heading 1

## Heading 2

---

Pandoc will first try to download the images that are referenced in the markdown file. If the website is protected against scraping, the images won't download. 

When Pandoc converts this file to Word, it will search for the image file locally in the input folder to insert it into the Word document. You should see a blue rectangle and a smile face in the Word document:

![Alt text: Smiley Blue Face](https://www.google.com/pretend/test-simple.png)

Hopefully that worked!

---

But the image "nonexistent.png" doesn't exist locally (in the input folder). So in the Word document this image will simply display as a clickable link to the online image:

![Image on Midjourney](https://cdn.document360.io/3040c2b6-fead-4744-a3a9-d56d621c6c7e/Images/Documentation/MJ_ImagePrompt_Statue_Flowers.jpg)

---

I wonder how Pandoc converts a python code blocks into Word.

```python
print("Hello World")
result = 5 + 5
print(f"5 + 5 = {result}")
# This is the last line of the code block.
```

**THE END.**