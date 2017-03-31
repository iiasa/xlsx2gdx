xlsx2gdx bash script to convert xslx to GAMS gdx 
======

### Dependencies 

1. *xlsx2csv*: convert xslx to csv (see [http://github.com/dilshod/xlsx2csv](http://github.com/dilshod/xlsx2csv)). To install use `sudo easy_install xlsx2csv` or `pip install xlsx2csv`

2. *GNU sed*: sed (stream editor) is a non-interactive command-line text editor (see [https://www.gnu.org/software/sed/](https://www.gnu.org/software/sed/))

3. *gams*: (see [https://www.gams.com/](https://www.gams.com/))

### Installation

1. Clone the repository `git clone https://github.com/iiasa/xlsx2gdx.git`

2. Move to the project folder `cd xlsx2gdx`

3. Run `./install.sh` 

### Usage

`gdxxrw Inputfile=./data/example.xlsx Outputfile=./data/example.gdx Par=paramter Rng='Sheet1!b4:g17' Rdim=2 Cdim=2`

