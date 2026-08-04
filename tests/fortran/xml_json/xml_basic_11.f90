! vybe-test: fortran/xml_json/xml_basic_11
! origin: languages/fortran/tests/fortran/test_xml_json.rs
program p
implicit none
character(len=64) :: s
s = '<a/>'
print *, s
end program p
