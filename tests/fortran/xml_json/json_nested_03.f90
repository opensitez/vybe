! vybe-test: fortran/xml_json/json_nested_03
! origin: languages/fortran/tests/fortran/test_xml_json.rs
program p
implicit none
character(len=128) :: s
s = '{"a":{"b":2}}'
print *, s
end program p
