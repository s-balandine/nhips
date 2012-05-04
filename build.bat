set epic="C:\Program Files (x86)\Epidata\epic.exe"

cd /d src\epidata

echo . > "..\..\build\field interviewer.rec"
echo . > "..\..\build\field supervisor.rec"
echo . > "..\..\build\household.rec"
echo . > "..\..\build\household member.rec"
echo . > "..\..\build\office editor.rec"
echo . > "..\..\build\office keyer.rec"
echo . > "..\..\build\eligible man.rec"
echo . > "..\..\build\eligible woman.rec"

%epic% rev "field interviewer.qes"  "..\..\build\field interviewer.rec" AUTO FORCE
%epic% rev "field supervisor.qes"   "..\..\build\field supervisor.rec"  AUTO FORCE
%epic% rev "household.qes"          "..\..\build\household.rec"         AUTO FORCE
%epic% rev "household member.qes"   "..\..\build\household member.rec"  AUTO FORCE
%epic% rev "office editor.qes"      "..\..\build\office editor.rec"     AUTO FORCE
%epic% rev "office keyer.qes"       "..\..\build\office keyer.rec"      AUTO FORCE
%epic% rev "eligible man.qes"       "..\..\build\eligible man.rec"      AUTO FORCE
%epic% rev "eligible woman.qes"     "..\..\build\eligible woman.rec"    AUTO FORCE

%epic% import TXT "persons.csv" "..\..\build\field interviewer.rec" delim=; q=all replace
%epic% import TXT "persons.csv" "..\..\build\field supervisor.rec"  delim=; q=all replace
%epic% import TXT "persons.csv" "..\..\build\office editor.rec"     delim=; q=all replace
%epic% import TXT "persons.csv" "..\..\build\office keyer.rec"      delim=; q=all replace

cd ..\..\build

del *.old.rec

pause