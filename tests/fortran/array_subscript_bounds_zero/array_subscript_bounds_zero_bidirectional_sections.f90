! vybe-test: fortran/array_subscript_bounds_zero/array_subscript_bounds_zero_bidirectional_sections
! origin: languages/fortran/tests/fortran/test_array_subscript_bounds_zero.rs

program array_subscript_bounds_zero_bidirectional_sections
    integer :: values(0:9)
    values = (/1, 2, 3, 4, 5, 6, 7, 8, 9, 10/)
    if ((values(2:8:2)(1)) /= 2) then
    print *, "FAIL: want [2] got [", values(2:8:2)(1), "]"
    stop 1
end if
    if ((values(8:2:-3)(2)) /= 8) then
    print *, "FAIL: want [8] got [", values(8:2:-3)(2), "]"
    stop 1
end if
    if ((size(values(2:8:2))) /= 4) then
    print *, "FAIL: want [4] got [", size(values(2:8:2)), "]"
    stop 1
end if
end program array_subscript_bounds_zero_bidirectional_sections
