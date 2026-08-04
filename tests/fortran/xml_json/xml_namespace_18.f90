! vybe-test: fortran/xml_json/xml_namespace_18
! origin: languages/fortran/tests/fortran/test_xml_json.rs
program p
implicit none
character(len=96) :: s
s = '<ns:a xmlns:ns="u"/>'
print *, s
end program p
