! vybe-test: fortran/array_implied_do_ordering/array_implied_do_ordering_stride_then_masked_fill
! origin: languages/fortran/tests/fortran/test_array_implied_do_ordering.rs

program array_implied_do_ordering_stride_then_masked_fill
    integer :: values(4)
    values = [(i, i = 2,8,2)]
    values = merge(values, 0, values > 2)
    if ((sum(values)) /= 18) then
    print *, "FAIL: want [18] got [", sum(values), "]"
    stop 1
end if
    if ((values(1)) /= 0) then
    print *, "FAIL: want [0] got [", values(1), "]"
    stop 1
end if
    if ((values(2)) /= 4) then
    print *, "FAIL: want [4] got [", values(2), "]"
    stop 1
end if
    if ((values(3)) /= 6) then
    print *, "FAIL: want [6] got [", values(3), "]"
    stop 1
end if
    if ((values(4)) /= 8) then
    print *, "FAIL: want [8] got [", values(4), "]"
    stop 1
end if
end program array_implied_do_ordering_stride_then_masked_fill
