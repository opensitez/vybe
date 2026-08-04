! vybe-test: fortran/array_subscript_bounds_zero/array_subscript_bounds_zero_lowered_indexing
! origin: languages/fortran/tests/fortran/test_array_subscript_bounds_zero.rs

program array_subscript_bounds_zero_lowered_indexing
    integer :: values(-2:2)
    integer :: i
    values = (/1, 2, 3, 4, 5/)
    i = values(-2) + values(-1) + values(0)
    if ((i) /= 6) then
    print *, "FAIL: want [6] got [", i, "]"
    stop 1
end if
    if ((values(2)) /= 4) then
    print *, "FAIL: want [4] got [", values(2), "]"
    stop 1
end if
end program array_subscript_bounds_zero_lowered_indexing
