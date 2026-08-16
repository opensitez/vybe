! vybe-test: fortran/array_constructor_repetition_counts/array_constructor_repetition_counts_fixed_shape_real_array
! origin: languages/fortran/tests/fortran/test_array_constructor_repetition_counts.rs

program array_constructor_repetition_counts_fixed_shape_real_array
    real :: values(5)
    integer :: n
    values = (/ (1.25, i = 1, 2), (0.75, i = 1, 3) /)
    n = size(values)
    if ((n) /= 5) then
    print *, "FAIL: want [5] got [", n, "]"
    stop 1
end if
    if ((nint(sum(values))) /= 5) then
    print *, "FAIL: want [5] got [", nint(sum(values)), "]"
    stop 1
end if
    if ((nint(values(1))) /= 1) then
    print *, "FAIL: want [1] got [", nint(values(1)), "]"
    stop 1
end if
    if ((nint(values(n))) /= 1) then
    print *, "FAIL: want [1] got [", nint(values(n)), "]"
    stop 1
end if
end program array_constructor_repetition_counts_fixed_shape_real_array
