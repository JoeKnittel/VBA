# VBA-Experiment

## Automated Vocabulary-Improvement App (2020)

*This project was completed as a fun coding experiment. More details can be found on my blog <a href = "https://joeknittel.github.io/2021/01/19/Extending-the-Functionality-of-Excel.html">here</a>.*

Everyone, from time time, whether while reading an old novel or perusing the news, encounters an obscure word and can't determine its meaning via context clues. A quick Google search normally does the job, but what happens when you inevitably forget the definition a few days down the road?

I sought to rectify this problem by creating an app in Excel that, just by typing in a word and hitting enter, would generate its definition, part of speech, and a link to its pronunciation. **With this file, each time I came across a new word, I would just type it in and save, creating an ever-growing list of tricky words' definitions that I could reference whenever I pleased, to help improve my vocabulary.**

In order to achieve the aforementioned automation task, I utilized VBA code (my own, and [external](https://github.com/VBA-tools/VBA-JSON) [modules](https://github.com/timhall/VBA-Dictionary) for parsing [JSON](https://www.json.org/json-en.html)) in conjunction with [Merriam Webster's API](https://dictionaryapi.com/products/api-collegiate-dictionary).

Below is an animated .gif demonstrating the file's general usage:

![](https://raw.githubusercontent.com/JosephKnittel/VBA/main/Images/vocab_demo.gif)

<hr>

### Contained in this repository: 

- Code written in VBA
- External VBA modules and classes 
- Underlying MS Excel file 
