! vybe-test: fortran/xml_json/json_payload_length_runtime
! origin: languages/fortran/tests/fortran/test_xml_json.rs

program p
    character(len=64) :: s
    s = '{"a":1}'
    if ((len_trim(s)) /= 7) then
    print *, "FAIL: want [7] got [", len_trim(s), "]"
    stop 1
end if
end program p
