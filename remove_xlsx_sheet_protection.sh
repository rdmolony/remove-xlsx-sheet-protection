echo 'Install sh requirements...'
sudo apt-get install zip unzip

temp_dir=$(echo $1 | cut -f 1 -d '.')
echo '\nUnzipping '"$1"' to '"$temp_dir"' ...\n'
unzip $1 -d $temp_dir 

echo '\nRemoving sheet protection ...\n'
cd $temp_dir/xl/worksheets
sed -i 's/<sheetProtection [^>]*//g' *.xml

unlocked_file=$temp_dir'_unlocked.xlsx'
echo '\nSaving results as '"$unlocked_file"' ...'
cd ../..
find . -type f | xargs zip ../$unlocked_file
cd ..

echo '\nRemoving '"$temp_dir"' ...\n'
rm -fr $temp_dir
