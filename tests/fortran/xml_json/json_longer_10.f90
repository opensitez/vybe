! vybe-test: fortran/xml_json/json_longer_10
! origin: languages/fortran/tests/fortran/test_xml_json.rs
program p
implicit none
character(len=128) :: s
s = '{"items":[{"id":1},{"id":2}]}'
print *, s
end program p
