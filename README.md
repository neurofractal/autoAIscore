# autoAIscore

Code to automatically extract data from Autobiographical Interview (AI) transcripts, scored by hand in Microsoft Word.

## Installing Python

*Instructions to follow*

Required Packages: docx, lxml, zipfile, pandas, re

## Filbury Study

#### Marking .docx files

- One document should hold transcriptions for all events
- Events should be clearly separated by Event XX before each transcription of the event
- For each event, split details into new paragraphs (new line)
- For each detail you wish to score, highlight the relevent text and add a new comment

- The comment should contain 4 letters, corresponding to the (adapted) Autobiographical Interview scoring protocol:

>- Letter 1: I (Internal Detail) or E (External Detail)
>- Letter 2-3: EV (Event), PE (Perceptual), TI (Time), PL (Place), TH (Thought/Emotion), SE (Semantic), RE (Repetition), OT (Other)
>- Letter 4: T (True), F (False), U (Unverifiable)

It should look like this:

![](./media/example_filbury1.png)

#### Running the code

In the command line run:

```python
python3 process_AIscores_filbury.py path_to_docxfile
```

The script should print any warnings/errors encountered, and export a .csv file with the same name as the .docx file, organised like so:

![](./media/example_filbury2.png)


## Conventional AI

*Instructions to follow*
