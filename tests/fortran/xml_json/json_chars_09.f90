! vybe-test: fortran/xml_json/json_chars_09
! origin: languages/fortran/tests/fortran/test_xml_json.rs
program p
implicit none
character(len=64) :: s
s = '{"name":"fortran"}'
print *, s
end program p
