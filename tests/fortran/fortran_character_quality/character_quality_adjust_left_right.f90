! vybe-test: fortran/fortran_character_quality/character_quality_adjust_left_right
! origin: languages/fortran/tests/fortran/test_fortran_character_quality.rs

program character_quality_adjust_left_right
    character(len=10) :: left
    character(len=10) :: right
    left = adjustl('  hello')
    right = adjustr('hello  ')
    if (trim(trim(left)) /= "hello") then
    print *, "FAIL: want [hello] got [", trim(left), "]"
    stop 1
end if
    if (trim(trim(right)) /= "hello") then
    print *, "FAIL: want [hello] got [", trim(right), "]"
    stop 1
end if
end program character_quality_adjust_left_right
