! vybe-test: fortran/array_constructor_repetition_counts/array_constructor_repetition_counts_fixed_shape_integer_array
! origin: languages/fortran/tests/fortran/test_array_constructor_repetition_counts.rs

program array_constructor_repetition_counts_fixed_shape_integer_array
    integer :: values(6)
    values = (/ 3 * 4, 2 * 6, 1 * 10 /)
    if ((size(values)) /= 6) then
    print *, "FAIL: want [6] got [", size(values), "]"
    stop 1
end if
    if ((values(1)) /= 4) then
    print *, "FAIL: want [4] got [", values(1), "]"
    stop 1
end if
    if ((values(4)) /= 6) then
    print *, "FAIL: want [6] got [", values(4), "]"
    stop 1
end if
    if ((sum(values)) /= 34) then
    print *, "FAIL: want [34] got [", sum(values), "]"
    stop 1
end if
end program array_constructor_repetition_counts_fixed_shape_integer_array
