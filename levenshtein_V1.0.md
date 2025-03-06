# Levenshtein formula V1.0

## Description

The Levenshtein formula calculates the Levenshtein distance between two strings, measuring the minimum number of single-character edits (insertions, deletions, or substitutions) required to transform one string into the other.

*Applies only to Microsoft 365*

## Arguments

| Arguments | Type | Description | Values |
|:--------------|:-------:|:-------------------------------------------------------------|:--------------------------------|
| firstWord     | Variant | First element to compare.                                     |  |
| nextWord      | Variant | Second element to compare.                                    |  |
| caseSensitive | Bool    | Defines if the Levenshtein comparison is case-sensitive.    | `TRUE`/`t` <br> `FALSE`/`f` <br> **OMITTED**: `TRUE` |
| outputType    | Str     | Defines the output type: a numerical integer (count of differences) or percentage of difference (based on the longer word). | `percent`/`p`/`%` <br> `integer`/`i`/`numeric`/`number` <br> **OMITTED**: `i` |
| printTable    | Str     | Defines the format of the output: the entire table (`true`), only the numerical part of the table (`numeric only`), or the Levenshtein value (`false`). | `true`/`t` <br> `false`/`f` <br> `numeric only`/`no` <br> **OMITTED**: `f` |

## Formula
```vb
=LAMBDA(firstWord,nextWord,[caseSensitive],[outputType],[printTable],LET(
caseSens,IF(ISOMITTED(caseSensitive),TRUE,IFERROR(SWITCH(caseSensitive,TRUE,TRUE,"t",TRUE,FALSE,FALSE,"f",FALSE),TRUE)),
outType,IF(ISOMITTED(outputType),"i",IFERROR(SWITCH(LOWER(outputType),"p","p","percent","p","%","p","i","i","int","i","integer","i","numeric","i","number","i"),"i")),
printTbl,IF(ISOMITTED(printTable),"f",IFERROR(SWITCH(LOWER(printTable),"true","t","t","t","false","f","f","f","no","no","numeric only","no"),"f")),

nbRow,LEN(firstWord)+1,
nbCol,LEN(nextWord)+1,
seqFirstWord,LAMBDA(tbl,IF(caseSens,tbl,LOWER(tbl)))(MAKEARRAY(nbRow-1,1,LAMBDA(r,c,RIGHT(LEFT(firstWord,r),1)))),
seqNextWord,LAMBDA(tbl,IF(caseSens,tbl,LOWER(tbl)))(MAKEARRAY(1,LEN(nextWord),LAMBDA(r,c,RIGHT(LEFT(nextWord,c),1)))),
firstRow,MAKEARRAY(1,nbCol,LAMBDA(r,c,c-1)),

newLine,LAMBDA(tbl,rowNum,LET(
rowAbove,CHOOSEROWS(tbl,rowNum),
DROP(REDUCE("",SEQUENCE(nbCol),LAMBDA(v,c,HSTACK(v,IF(c=1,INDEX(rowAbove,1,1)+1,IF(EXACT(INDEX(seqFirstWord,rowNum),INDEX(seqNextWord,1,c-1)),INDEX(rowAbove,1,c-1),1+MIN(INDEX(v,1,c),INDEX(rowAbove,1,c),INDEX(rowAbove,1,c-1))))))),0,1))),

formatMultiplier,SWITCH(outType,"i",1,"p",100/(MAX(nbRow,nbCol)-1)),
levenshteinMatrix,formatMultiplier*REDUCE(firstRow,SEQUENCE(nbRow-1),LAMBDA(t,r,VSTACK(t,newLine(t,r)))),
formatOutput,SWITCH(printTbl,"t",HSTACK(VSTACK("*","",seqFirstWord),VSTACK(HSTACK("",seqNextWord),levenshteinMatrix)),"no",levenshteinMatrix,"f",INDEX(levenshteinMatrix,nbRow,nbCol)),

formatOutput))
```

## Example of Usage

### Inputs
- **First_Word:** `Sitten`
- **Second_Word:** `Kitten`
- **levenshtein_V1:** Name given to the formula in Excel name manager.

### Example 1

```vb
=levenshtein_V1(First_Word,Second_Word,"f","p","t")
```

|*||k|i|t|t|e|n|
|:---:|:---:|:---:|:---:|:---:|:---:|:---:|:---:|
||0|16,7|33,3|50|66,7|83,3|100|
|s|16,7|16,7|33,3|50|66,7|83,3|100|
|i|33,3|33,3|16,7|33,3|50|66,7|83,3|
|t|50|50|33,3|16,7|33,3|50|66,7|
|t|66,7|66,7|50|33,3|16,7|33,3|50|
|e|83,3|83,3|66,7|50|33,3|16,7|33,3|
|n|100|100|83,3|66,7|50|33,3|16,7|

### Example 2

```vb
=levenshtein_V1(First_Word,Second_Word,"f","i","t")
```

|*||k|i|t|t|e|n|
|:---:|:---:|:---:|:---:|:---:|:---:|:---:|:---:|
||0|1|2|3|4|5|6|
|s|1|1|2|3|4|5|6|
|i|2|2|1|2|3|4|5|
|t|3|3|2|1|2|3|4|
|t|4|4|3|2|1|2|3|
|e|5|5|4|3|2|1|2|
|n|6|6|5|4|3|2|1|

### Example 3

```vb
=levenshtein_V1(First_Word,Second_Word,,,"t")
```

|A|B|C|D|E|F|G|
|:---:|:---:|:---:|:---:|:---:|:---:|:---:|
|0|1|2|3|4|5|6|
|1|1|2|3|4|5|6|
|2|2|1|2|3|4|5|
|3|3|2|1|2|3|4|
|4|4|3|2|1|2|3|
|5|5|4|3|2|1|2|
|6|6|5|4|3|2|1|

### Example 4

```vb
=levenshtein_V1(First_Word,Second_Word)
```

`1`
