! vybe-test: fortran/xml_json/xml_and_json_concat_runtime
! origin: languages/fortran/tests/fortran/test_xml_json.rs

    program p
    character(len=128) :: xml
    character(len=128) :: json
    character(len=256) :: combined
    xml = '<row><id>1</id></row>'
    json = '{"id":1}'
    combined = trim(xml) // trim(json)
    if ((index(combined, '{')) /= 22) then
    print *, "FAIL: want [22] got [", index(combined, '{'), "]"
    stop 1
end if
    if (trim(trim(combined(1:5))) /= "<row>") then
    print *, "FAIL: want [<row>] got [", trim(combined(1:5)), "]"
    stop 1
end if
    if (trim(trim(combined(22:25))) /= "{'id") then
    print *, "FAIL: want [{'id] got [", trim(combined(22:25)), "]"
    stop 1
end if
end program p
