! vybe-test: fortran/array_subscript_bounds_zero/array_subscript_bounds_zero_char_length_preserved
! origin: languages/fortran/tests/fortran/test_array_subscript_bounds_zero.rs

program array_subscript_bounds_zero_char_length_preserved
    character(len=3) :: items(0:3)
    items = (/'aaa', 'bbb', 'ccc', 'ddd'/)
    if (trim(trim(items(0))) /= "aaa") then
    print *, "FAIL: want [aaa] got [", trim(items(0)), "]"
    stop 1
end if
    if (trim(trim(items(3))) /= "ddd") then
    print *, "FAIL: want [ddd] got [", trim(items(3)), "]"
    stop 1
end if
    if ((size(items)) /= 4) then
    print *, "FAIL: want [4] got [", size(items), "]"
    stop 1
end if
end program array_subscript_bounds_zero_char_length_preserved
