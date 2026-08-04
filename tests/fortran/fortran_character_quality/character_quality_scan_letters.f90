! vybe-test: fortran/fortran_character_quality/character_quality_scan_letters
! origin: languages/fortran/tests/fortran/test_fortran_character_quality.rs

program character_quality_scan_letters
    character(len=20) :: source
    source = '0012ab34'
    if ((scan(source, 'ab')) /= 5) then
    print *, "FAIL: want [5] got [", scan(source, 'ab'), "]"
    stop 1
end if
    if ((scan(source, 'xyz')) /= 0) then
    print *, "FAIL: want [0] got [", scan(source, 'xyz'), "]"
    stop 1
end if
end program character_quality_scan_letters
