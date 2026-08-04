! vybe-test: fortran/fortran_character_quality/character_quality_character_to_int_parse
! origin: languages/fortran/tests/fortran/test_fortran_character_quality.rs

program character_quality_character_to_int_parse
    character(len=8) :: token
    integer :: value
    token = '007'
    read (token, '(I0)') value
    if ((value + 1) /= 8) then
    print *, "FAIL: want [8] got [", value + 1, "]"
    stop 1
end if
end program character_quality_character_to_int_parse
