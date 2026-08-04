! vybe-test: fortran/xml_json/xml_json_mix_20
! origin: languages/fortran/tests/fortran/test_xml_json.rs
program p
implicit none
character(len=160) :: a, b
a = '<row><id>1</id></row>'
b = '{"id":1}'
print *, a
print *, b
end program p
