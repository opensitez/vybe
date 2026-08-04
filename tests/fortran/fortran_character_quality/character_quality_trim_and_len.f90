! vybe-test: fortran/fortran_character_quality/character_quality_trim_and_len
! origin: languages/fortran/tests/fortran/test_fortran_character_quality.rs

program character_quality_trim_and_len
    character(len=20) :: text
    text = 'fortran   '
    if ((len(text)) /= 20) then
    print *, "FAIL: want [20] got [", len(text), "]"
    stop 1
end if
    if ((len_trim(text)) /= 7) then
    print *, "FAIL: want [7] got [", len_trim(text), "]"
    stop 1
end if
    if ((len_trim(trim(text))) /= 7) then
    print *, "FAIL: want [7] got [", len_trim(trim(text)), "]"
    stop 1
end if
end program character_quality_trim_and_len
