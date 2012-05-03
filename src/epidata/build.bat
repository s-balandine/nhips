set epic="C:\Program Files (x86)\Epidata\epic.exe"

echo . > "field interviewer.rec"
echo . > "field supervisor.rec"
echo . > "household.rec"
echo . > "household member.rec"
echo . > "office editor.rec"
echo . > "office keyer.rec"
echo . > "eligible man.rec"
echo . > "eligible woman.rec"

%epic% rev "field interviewer.qes" * FIRST FORCE
%epic% rev "field supervisor.qes" * FIRST FORCE
%epic% rev "household.qes" * FIRST FORCE
%epic% rev "household member.qes" * FIRST FORCE
%epic% rev "office editor.qes" * FIRST FORCE
%epic% rev "office keyer.qes" * FIRST FORCE
%epic% rev "eligible man.qes" * FIRST FORCE
%epic% rev "eligible woman.qes" * FIRST FORCE

pause