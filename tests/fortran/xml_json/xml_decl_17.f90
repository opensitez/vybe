! vybe-test: fortran/xml_json/xml_decl_17
! origin: languages/fortran/tests/fortran/test_xml_json.rs
program p
implicit none
character(len=96) :: s
s = '<?xml version="1.0"?><a/>'
print *, s
end program p
