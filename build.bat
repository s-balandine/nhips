set epic="C:\Program Files (x86)\Epidata\epic.exe"

cd /d src\epidata

echo . > "field interviewer.rec"
echo . > "field supervisor.rec"
echo . > "household.rec"
echo . > "household member.rec"
echo . > "office editor.rec"
echo . > "office keyer.rec"
echo . > "eligible man.rec"
echo . > "eligible woman.rec"

%epic% rev "field interviewer.qes"  "field interviewer.rec" AUTO FORCE
%epic% rev "field supervisor.qes"   "field supervisor.rec"  AUTO FORCE
%epic% rev "household.qes"          "household.rec"         AUTO FORCE
%epic% rev "household member.qes"   "household member.rec"  AUTO FORCE
%epic% rev "office editor.qes"      "office editor.rec"     AUTO FORCE
%epic% rev "office keyer.qes"       "office keyer.rec"      AUTO FORCE
%epic% rev "eligible man.qes"       "eligible man.rec"      AUTO FORCE
%epic% rev "eligible woman.qes"     "eligible woman.rec"    AUTO FORCE

%epic% import TXT "person.csv" "field interviewer.rec" qes="field interviewer.rec" delim=; q=none replace ignorefirst
%epic% import TXT "person.csv" "field supervisor.rec"  qes="field supervisor.rec"  delim=; q=none replace ignorefirst
%epic% import TXT "person.csv" "office editor.rec"     qes="office editor.rec"     delim=; q=none replace ignorefirst
%epic% import TXT "person.csv" "office keyer.rec"      qes="office keyer.rec"      delim=; q=none replace ignorefirst

del  *.old.rec
copy *.chk ..\..\build
move *.rec ..\..\build



pause