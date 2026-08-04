! vybe-test: fortran/fortran_character_quality/character_quality_repeat_fill
! origin: languages/fortran/tests/fortran/test_fortran_character_quality.rs

program character_quality_repeat_fill
    character(len=20) :: text
    text = repeat('a', 5)
    if ((len_trim(text)) /= 5) then
    print *, "FAIL: want [5] got [", len_trim(text), "]"
    stop 1
end if
    if (trim(text) /= "aaaaa") then
    print *, "FAIL: want [aaaaa] got [", text, "]"
    stop 1
end if
end program character_quality_repeat_fill
