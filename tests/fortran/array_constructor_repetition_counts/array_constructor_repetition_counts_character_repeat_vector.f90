! vybe-test: fortran/array_constructor_repetition_counts/array_constructor_repetition_counts_character_repeat_vector
! origin: languages/fortran/tests/fortran/test_array_constructor_repetition_counts.rs

program array_constructor_repetition_counts_character_repeat_vector
    character(len=4), allocatable :: values(:)
    values = (/ 3 * 'a', 2 * 'xy' /)
    if ((size(values)) /= 5) then
    print *, "FAIL: want [5] got [", size(values), "]"
    stop 1
end if
    if ((len(values(1))) /= 4) then
    print *, "FAIL: want [4] got [", len(values(1)), "]"
    stop 1
end if
    if ((len_trim(values(2))) /= 1) then
    print *, "FAIL: want [1] got [", len_trim(values(2)), "]"
    stop 1
end if
    if ((len_trim(values(4))) /= 2) then
    print *, "FAIL: want [2] got [", len_trim(values(4)), "]"
    stop 1
end if
end program array_constructor_repetition_counts_character_repeat_vector
