! vybe-test: fortran/fortran_character_quality/character_quality_substring_bounds
! origin: languages/fortran/tests/fortran/test_fortran_character_quality.rs

program character_quality_substring_bounds
    character(len=20) :: text
    text = 'runtime-check'
    if (trim(text(1:7)) /= "runtime") then
    print *, "FAIL: want [runtime] got [", text(1:7), "]"
    stop 1
end if
    if (trim(text(9:13)) /= "h-che") then
    print *, "FAIL: want [h-che] got [", text(9:13), "]"
    stop 1
end if
end program character_quality_substring_bounds
