! vybe-test: fortran/xml_json/json_number_07
! origin: languages/fortran/tests/fortran/test_xml_json.rs
program p
implicit none
character(len=64) :: s
s = '{"n":12.5}'
print *, s
end program p
