! vybe-test: fortran/xml_json/json_key_presence_runtime
! origin: languages/fortran/tests/fortran/test_xml_json.rs

program p
    character(len=64) :: s
    s = '{"user":"fortran"}'
    if ((index(s, '"user"')) /= 2) then
    print *, "FAIL: want [2] got [", index(s, '"user"'), "]"
    stop 1
end if
    if ((index(s, 'fortran')) /= 10) then
    print *, "FAIL: want [10] got [", index(s, 'fortran'), "]"
    stop 1
end if
end program p
