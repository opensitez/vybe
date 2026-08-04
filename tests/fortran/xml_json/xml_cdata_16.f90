! vybe-test: fortran/xml_json/xml_cdata_16
! origin: languages/fortran/tests/fortran/test_xml_json.rs
program p
implicit none
character(len=96) :: s
s = '<a><![CDATA[x]]></a>'
print *, s
end program p
