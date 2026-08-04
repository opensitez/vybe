! vybe-test: fortran/xml_json/xml_tag_scan_runtime
! origin: languages/fortran/tests/fortran/test_xml_json.rs

    program p
    character(len=64) :: s
    integer :: i
    s = '<msg><id>7</id></msg>'
    i = index(s, '<id>')
    if ((i) /= 6) then
    print *, "FAIL: want [6] got [", i, "]"
    stop 1
end if
    if ((len_trim(trim(s(i:)))) /= 16) then
    print *, "FAIL: want [16] got [", len_trim(trim(s(i:))), "]"
    stop 1
end if
    if ((index(s(i:), '</id>')) /= 6) then
    print *, "FAIL: want [6] got [", index(s(i:), '</id>'), "]"
    stop 1
end if
end program p
