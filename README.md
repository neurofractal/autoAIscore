# autoAIscore
[![made-with-python](https://img.shields.io/badge/Made%20with-Python-1f425f.svg)](https://www.python.org/)

Code to automatically extract data from Autobiographical Interview (AI) transcripts, scored by hand in Microsoft Word.


## Citation

[![DOI](https://zenodo.org/badge/429506915.svg)](https://zenodo.org/badge/latestdoi/429506915) 

**Robert Seymour. (2022).** neurofractal/autoAIscore: v0.11 (v0.11). *Zenodo.* https://doi.org/10.5281/zenodo.7225874

## Installing Python

1. Download and Install Anaconda Python from **[HERE](https://www.anaconda.com/products/individual)**
2. Download the autoAIscore repository in Documents (or wherever you want it)
3. Open Anaconda3 --> Anaconda Prompt
4. In the terminal type:

```bash
pip install python-docx lxml 
```
5. In the terminal type:

```bash
# Change path to match your Documents path
cd /users/XXX/Documents/autoAIscore
```

6. To test whether the script(s) are working, in the terminal type:

```bash
# Change paths to match your data
python process_AIscores /users/XXX/Documents/autoAIscore/test_AI /users/XXX/Documents/autoAIscore/test_AI/out.csv
```

*!! Email Rob or IT Support if you have any issues !!*

## Filbury Study

#### Marking .docx files

- One document should hold transcriptions each participant and for all events
- Events should be clearly separated by 'Event XX' before each transcription of the event
- For each event, split details into new paragraphs (new line) according to AI scoring protocols
- For each detail you wish to score, highlight the relevent text and add a new comment
> Speed this up by creating a keyboard shortcut for 'New Comment'

- The comment should contain a 4 letter code, corresponding to the (adapted) Autobiographical Interview scoring protocol:

>- Letter 1: I (Internal Detail) or E (External Detail)
>- Letter 2-3: EV (Event), PE (Perceptual), TI (Time), PL (Place), TH (Thought/Emotion), SE (Semantic), RE (Repetition), OT (Other)
>- Letter 4: T (True), F (False), U (Unverifiable)

It should look like this:

![](./media/example_filbury1.png)

#### Running the code

In the command line run (for example):

```python
python3 process_AIscores_filbury.py test_filbury/test_participant01.docx
```

The script should print any warnings/errors encountered, and export a .csv file with the same name as the .docx file, organised like so:

![](./media/example_filbury2.png)


## Conventional AI

#### Marking .docx files

- Each folder should contain separate transcription documents for each memory, from one participant and **NO OTHER .docx FILES**
- Split details into new paragraphs (new line) according to AI scoring protocol
- For each detail you wish to score, highlight the relevent text and add a new comment
> Speed this up by creating a keyboard shortcut for 'New Comment'

- The comment should contain a 3 letter code, corresponding to the Autobiographical Interview scoring protocol:

>- Letter 1: I (Internal Detail) or E (External Detail)
>- Letter 2-3: EV (Event), PE (Perceptual), TI (Time), PL (Place), TH (Thought/Emotion), SE (Semantic), RE (Repetition), OT (Other)

It should look like this:

![](./media/example_AI1.png)

#### Running the code

In the command line run (for example):

```python
python3 process_AIscores.py test_AI/ testAI/out.csv
```

The script should print any warnings/errors encountered, and export a .csv file organised like so:

![](./media/example_AI2.png)





