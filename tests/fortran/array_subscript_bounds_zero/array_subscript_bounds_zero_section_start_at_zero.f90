! vybe-test: fortran/array_subscript_bounds_zero/array_subscript_bounds_zero_section_start_at_zero
! origin: languages/fortran/tests/fortran/test_array_subscript_bounds_zero.rs

program array_subscript_bounds_zero_section_start_at_zero
    integer :: values(0:5)
    values = (/0, 1, 2, 3, 4, 5/)
    if ((sum(values(0:3))) /= 3) then
    print *, "FAIL: want [3] got [", sum(values(0:3)), "]"
    stop 1
end if
    if ((size(values(0:3))) /= 4) then
    print *, "FAIL: want [4] got [", size(values(0:3)), "]"
    stop 1
end if
    if ((values(0:3)(1)) /= 0) then
    print *, "FAIL: want [0] got [", values(0:3)(1), "]"
    stop 1
end if
    if ((values(0:3)(4)) /= 3) then
    print *, "FAIL: want [3] got [", values(0:3)(4), "]"
    stop 1
end if
end program array_subscript_bounds_zero_section_start_at_zero
