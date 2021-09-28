#!/bin/bash

CNTLNAME=""
TESTNAME=""

CNTLDIR=""
TESTDIR=""
CMPDIR=""
CNTLDIR2=""
TESTDIR2=""
CMPDIR2=""

ANL="" 


PYTHON37=""
ROOTDIR=""
WKROOTDIR=""
mkdir ${WKROOTDIR}

sed '{
     s|@@@ANL@@@|'${ANL}'|g
     s|@@@ANL,,@@@|'${ANL,,}'|g
     s|@@@TESTNAME@@@|'${TESTNAME}'|g
     s|@@@CNTLNAME@@@|'${CNTLNAME}'|g
     s|@@@TESTDIR@@@|'${TESTDIR}'|g
     s|@@@CNTLDIR@@@|'${CNTLDIR}'|g
     s|@@@CMPDIR@@@|'${CMPDIR}'|g
     s|@@@TESTDIR2@@@|'${TESTDIR2}'|g
     s|@@@CNTLDIR2@@@|'${CNTLDIR2}'|g
     s|@@@CMPDIR2@@@|'${CMPDIR2}'|g
     s|[[:blank:]]*$||g
     }' ${ROOTDIR}/tables/table.txt > ${WKROOTDIR}/table.txt

mkdir -p "${WKROOTDIR}/slides"
cp ${ROOTDIR}/tables/padding.txt "${WKROOTDIR}/slides/padding.txt"

cd ${WKROOTDIR}

while read line; do
  numfile=`echo ${line} | wc -l`
  if [[ x${line} == x ]]; then
    continue 
  elif [[ ${line} == slide* ]]; then
    slide=`echo ${line} | xargs`   # remove whitespaces from string
    mkdir -p "${WKROOTDIR}/slides/${slide}"
    cd "${WKROOTDIR}/slides/${slide}"
  elif [[ ${numfile} == 1 ]]; then
    read header partial_path dire <<< $line
    fig_path=`find ${dire} -path *${partial_path}`
    echo "fig_path=${fig_path}"
    figname=`basename ${fig_path}`
    figname="${header}_${figname}"
    if [[ ${figname} == ${header}_ ]]; then
      cp ${ROOTDIR}/images/no_image.png "${WKROOTDIR}/slides/${slide}/${header}no_image.png"
    else
      cp $fig_path "${WKROOTDIR}/slides/${slide}/${figname}"
    fi
  else
    echo "multiple files hit for ${fig_path}!!"
    cp ${ROOTDIR}/images/no_image.png "${WKROOTDIR}/slides/${slide}/${header}no_image.png"
  fi
done <"${WKROOTDIR}/table.txt"

${PYTHON37} "${ROOTDIR}/insertFigPPT.py" "${WKROOTDIR}/slides"

cp "${WKROOTDIR}/slides/temp.pptx" ${ROOTDIR}

cd ..
rm -rf ${WKROOTDIR}


