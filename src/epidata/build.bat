set epic="C:\Program Files (x86)\Epidata\epic.exe"

echo . > "field interviewer.rec"
echo . > "field supervisor.rec"
echo . > "household.rec"
echo . > "household member.rec"
echo . > "office editor.rec"
echo . > "office keyer.rec"
echo . > "eligible man.rec"
echo . > "eligible woman.rec"

%epic% rev "field interviewer.qes" * AUTO FORCE
%epic% rev "field supervisor.qes" * AUTO FORCE
%epic% rev "household.qes" * AUTO FORCE
%epic% rev "household member.qes" * AUTO FORCE
%epic% rev "office editor.qes" * AUTO FORCE
%epic% rev "office keyer.qes" * AUTO FORCE
%epic% rev "eligible man.qes" * AUTO FORCE
%epic% rev "eligible woman.qes" * AUTO FORCE

%epic% import TXT "persons.csv" "field interviewer.rec" delim=; q=none replace
%epic% import TXT "persons.csv" "field supervisor.rec" delim=; q=none replace
%epic% import TXT "persons.csv" "office editor.rec" delim=; q=none replace
%epic% import TXT "persons.csv" "office keyer.rec" delim=; q=none replace


del *.old.rec

pause