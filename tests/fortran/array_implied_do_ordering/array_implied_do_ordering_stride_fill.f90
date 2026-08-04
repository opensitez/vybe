! vybe-test: fortran/array_implied_do_ordering/array_implied_do_ordering_stride_fill
! origin: languages/fortran/tests/fortran/test_array_implied_do_ordering.rs

program array_implied_do_ordering_stride_fill
    integer :: values(4)
    values = [(i, i=1,8,2)]
    if ((values(1)) /= 1) then
    print *, "FAIL: want [1] got [", values(1), "]"
    stop 1
end if
    if ((values(4)) /= 7) then
    print *, "FAIL: want [7] got [", values(4), "]"
    stop 1
end if
    if ((sum(values)) /= 16) then
    print *, "FAIL: want [16] got [", sum(values), "]"
    stop 1
end if
end program array_implied_do_ordering_stride_fill
