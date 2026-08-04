! vybe-test: fortran/xml_json/xml_text_15
! origin: languages/fortran/tests/fortran/test_xml_json.rs
program p
implicit none
character(len=64) :: s
s = '<msg>hello</msg>'
print *, s
end program p
