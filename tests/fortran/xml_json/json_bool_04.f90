! vybe-test: fortran/xml_json/json_bool_04
! origin: languages/fortran/tests/fortran/test_xml_json.rs
program p
implicit none
character(len=32) :: s
s = '{"ok":true}'
print *, s
end program p
