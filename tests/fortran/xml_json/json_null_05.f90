! vybe-test: fortran/xml_json/json_null_05
! origin: languages/fortran/tests/fortran/test_xml_json.rs
program p
implicit none
character(len=32) :: s
s = '{"x":null}'
print *, s
end program p
