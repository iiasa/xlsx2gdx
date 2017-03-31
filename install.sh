#!/bin/bash
#title           :install.sh
#description     :Install gdxxrw for unix
#author          :Victor Maus
#date            :2017-03-30
#usage           :./config.sh 
#=================================================================================

echo "Checking dependencies"

# Check xlsx2csv - xlsx to csv converter (http://github.com/dilshod/xlsx2csv)
command -v xlsx2csv >/dev/null 2>&1 || { echo >&2 "Command xlsx2csv not found. See http://github.com/dilshod/xlsx2csv. To install use sudo easy_install xlsx2csv or pip install xlsx2csv."; exit 1; }
echo "xlsx2csv ... OK"

# Check sed 
command -v sed >/dev/null 2>&1 || { echo >&2 "Command sed not fount. See https://www.gnu.org/software/sed/."; exit 1; }
echo "sed ... OK" 

# Check gams 
gams_path=$(command -v gams | sed 's/\(.*\)\/gams/\1/') || { echo >&2 "Command gams not found. See https://www.gams.com/."; exit 1; }
echo "gams ... OK"

# Creating unix version of gdxxrw in gams installation folder 

cat >$gams_path/gdxxrw <<'EOL'
#!/bin/bash 
#title           :xlsx2gdx.sh
#description     :Convert xlsx file to gdx.
#author		 :Victor Maus
#date            :2017-03-24
#usage		 :./xlsx2gdx.sh inputfile=<xlsx file path> outputfile=<gdx file path> Par=<Par> Rng=<Rng> Rdim=<Rdim> Cdim=<Cdim>
#inputfile       :xlsx file path 
#outputfile      :gdx file path 
#Par             :GAMS_Parameter. Specify a GAMS parameter to be read from a GDX file and written to spreadsheet, or to be read from a spreadsheet and written to a GDX file.
#Rng             :Excel Range. The Excel Range for the data for the symbol. Note that an empty range is equivalent to the first cell of the first sheet.
#Rdim            :Row dimension: the number of columns in the data range that will be used to define the labels for the rows. The first Rdim columns of the data range will be used for the labels.
#Cdim            :Column dimension: the number of rows in the data range that will be used to define the labels for the columns. The first Cdim rows of the data range will be used for labels.
#=================================================================================

inputfile=$(echo $1 | cut -d '=' -f2) 
outputfile=$(echo $2 | cut -d '=' -f2) 
Par=$(echo $3 | cut -d '=' -f2) 
Rng=$(echo $4 | cut -d '=' -f2) 
Rdim=$(echo $5 | cut -d '=' -f2) 
Cdim=$(echo $6 | cut -d '=' -f2) 

[ "$inputfile" == "" ]  && { echo "Error: missing inputfile"; exit 1; }
[ "$outputfile" == "" ] && { echo "Error: missing outputfile"; exit 1; }
[ "$Par" == "" ]        && { echo "Error: missing Par"; exit 1; }
[ "$Rng" == "" ]        && { echo "Error: missing Rng"; exit 1; }
[ "$Rdim" == "" ]       && { echo "Error: missing Rdim"; exit 1; }
[ "$Cdim" == "" ]       && { echo "Error: missing Cdim"; exit 1; }

tmp_csv=$(date | sed -e 's/ /_/g' | sed -e 's/:/_/g').csv

# Get Exel sheet 
sheet_name=$(echo $Rng | cut -d '!' -f1)

# Get data bounding box
sheet_bbox=$(echo $Rng | cut -d '!' -f2)

# Get upper left corner 
sheet_ul=$(echo $sheet_bbox | cut -d ':' -f1)

# Get lower right corner 
if [ "$( echo $sheet_bbox | grep ':')" == "" ]; then
  sheet_lr=""
else
  sheet_lr=$(echo $sheet_bbox | cut -d ':' -f2)
fi

# Get initial row 
ri=${sheet_ul//[!0-9]/}

# Get final row 
rf=${sheet_lr//[!0-9]/}

# Get initial column
ci=${sheet_ul//[0-9]/}

# Get final column 
cf=${sheet_lr//[0-9]/}

# Columns to lower case  
ci=${ci,,}
cf=${cf,,}

# Columns to numbers 
L=({a..z})
N=${#L[@]}
MAX=16384
i=0
j=0
cni=0
cnf=0
aux=""

# Needs xlsx2csv - xlsx to csv converter (http://github.com/dilshod/xlsx2csv)
# sudo easy_install xlsx2csv  or  pip install xlsx2csv
# Crop col/row from original sheet 
xlsx2csv -d ',' -n ${sheet_name} ${inputfile} ${tmp_csv} 

#cat ${tmp_csv} > ${tmp_csv}_tmp3.csv
csv_nr=$(wc -l < ${tmp_csv})

if [ "$rf" == "" ]; then
  rf=$csv_nr
fi

while [  $i -lt $N ]; do
  if [ $cni -gt $MAX ]; then
    echo "Excel column '$ci' out of bounds"
    cni=""
    break
  fi
  col=$aux${L[$i]}
  let cni=cni+1
  let i=i+1
  if [ "$ci" == "$col" ]; then
    break
  elif [ $i -eq $N ]; then 
    aux=${L[$j]}
    let j=j+1
    i=0
  fi
done

if [ "$cf" == "" ]; then 
  cnf=$(head -1 ${tmp_csv} | sed 's/[^,]//g' | wc -c)
  #cnf="LastCol"
else
  i=0
  j=0
  while [  $i -lt $N ]; do
    if [ $cnf -gt $MAX ]; then
      echo "Excel column '$cf' out of bounds"
      cnf=""
      break
    fi
    col=$aux${L[$i]}
    let cnf=cnf+1
    let i=i+1
    if [ "$cf" == "$col" ]; then
      break
    elif [ $i -eq $N ]; then 
      aux=${L[$j]}
      let j=j+1
      i=0
    fi
  done
fi

# Crop col/row from original sheet
#echo rows: ${ri}-${rf} cols: ${cni}-${cnf}
sed -n ${ri},${rf}p ${tmp_csv} | cut -d , -f${cni}-${cnf} > ${tmp_csv}_tmp2.csv

# Melt csv 
let START_COL=Rdim+1
let START_ROW=Cdim+1
let END_COL=$(head -1 ${tmp_csv} | sed 's/[^,]//g' | wc -c)
let END_ROW=$(wc -l < ${tmp_csv})

# Index to new columns 
let dcols=Rdim+Cdim
let vcol=dcols+1
echo "$(printf '%*s' "$dcols" | tr ' ' ",")value" > ${tmp_csv}
dcols=1..$dcols

#echo rows: $START_ROW-$END_ROW cols: $START_COL-$END_COL

for ((j=START_COL; j<=END_COL; j++)); do
   new_col_dim=$(sed -n 1p ${tmp_csv}_tmp2.csv | cut -d , -f ${j})
   for ((k=2; k<=Cdim; k++)); do
      new_col_dim="$new_col_dim,$(sed -n ${k}p ${tmp_csv}_tmp2.csv | cut -d , -f ${j})"
   done
   for ((i=START_ROW; i<=END_ROW; i++)); do
     #echo New line: $j ----- $new_col_dim
     #echo "$new_col_dim,$(sed -n ${i}p ${tmp_csv}_tmp2.csv | cut -d , -f ${j})" 
     echo "$(sed -n ${i}p ${tmp_csv}_tmp2.csv | cut -d , -f 1-${Rdim}),$new_col_dim,$(sed -n ${i}p ${tmp_csv}_tmp2.csv | cut -d , -f ${j})" >> ${tmp_csv}
   done
done

# Creat GDX file 
csv2gdx FieldSep=Comma DecimalSep=Period Input=$tmp_csv ID=$Par UseHeader=y Index=\($dcols\) Value=\($vcol\) Output=$outputfile

# Remove tmp files 
rm -rf ${tmp_csv}
rm -rf ${tmp_csv}_tmp2.csv

echo Done

EOL

# Change gdxxrw permission to executable 
chmod +x $gams_path/gdxxrw

# Check if gdxxrw is executable 
if [[ ! -x "$gams_path/gdxxrw" ]]; then
    echo "File '$gams_path/gdxxrw' is not executable or found"
    exit 1 
fi

# Check gdxxrw command  
command -v gdxxrw >/dev/null 2>&1 || { echo >&2 "Command gdxxrw not fount. Something went wrong. Please check the dependencies and user permission of gams installation folder."; exit 1; }
echo "gdxxrw ... OK"

echo "Installation successfully completed!"



