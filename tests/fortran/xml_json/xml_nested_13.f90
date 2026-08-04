! vybe-test: fortran/xml_json/xml_nested_13
! origin: languages/fortran/tests/fortran/test_xml_json.rs
program p
implicit none
character(len=128) :: s
s = '<a><b>2</b></a>'
print *, s
end program p
