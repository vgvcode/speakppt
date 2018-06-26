#!/bin/bash
if [ $# -ne 1 ]; then
   echo "Usage:" $0 " /path-to-file"
   exit 1
fi
echo Translating text to speech...
python3 ./translate2speech.py $1
echo Done
for f in $1.*.wav
do
  echo Processing $f
  mv $f '/Users/vgvcode/Library/Application Scripts/com.microsoft.Powerpoint/.'
done
echo Translating text to Powerpoint slides...
python3 ./translate2ppt.py $1
echo done
echo Open your Powerpoint presentation $1.pptm
echo Click on Enable Macros
echo Go to Tools-Macro-Macros...
echo Run the PlaySlideShowAutoSlideAdvance macro
echo IMPORTANT:  Remember to save your ppt after it runs
