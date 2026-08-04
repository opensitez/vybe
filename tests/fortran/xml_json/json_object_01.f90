! vybe-test: fortran/xml_json/json_object_01
! origin: languages/fortran/tests/fortran/test_xml_json.rs
program p
implicit none
character(len=64) :: s
s = '{"a":1}'
print *, s
end program p
