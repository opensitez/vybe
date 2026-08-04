! vybe-test: fortran/xml_json/json_escape_06
! origin: languages/fortran/tests/fortran/test_xml_json.rs
program p
implicit none
character(len=64) :: s
s = '{"t":"a\nb"}'
print *, s
end program p
