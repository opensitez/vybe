! vybe-test: fortran/xml_json/json_array_02
! origin: languages/fortran/tests/fortran/test_xml_json.rs
program p
implicit none
character(len=64) :: s
s = '[1,2,3]'
print *, s
end program p
