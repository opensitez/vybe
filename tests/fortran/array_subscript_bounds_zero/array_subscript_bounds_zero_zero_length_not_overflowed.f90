! vybe-test: fortran/array_subscript_bounds_zero/array_subscript_bounds_zero_zero_length_not_overflowed
! origin: languages/fortran/tests/fortran/test_array_subscript_bounds_zero.rs

program array_subscript_bounds_zero_zero_length_not_overflowed
    integer :: values(0:-1)
    if ((size(values)) /= 0) then
    print *, "FAIL: want [0] got [", size(values), "]"
    stop 1
end if
end program array_subscript_bounds_zero_zero_length_not_overflowed
