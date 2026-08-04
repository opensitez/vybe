! vybe-test: fortran/array_implied_do_ordering/array_implied_do_ordering_implied_fill_for_real
! origin: languages/fortran/tests/fortran/test_array_implied_do_ordering.rs

program array_implied_do_ordering_implied_fill_for_real
    real, allocatable :: values(:)
    integer :: n
    values = [(1.0 * i / 2.0, i = 1,4)]
    n = nint(sum(values) * 10.0)
    if ((n) /= 50) then
    print *, "FAIL: want [50] got [", n, "]"
    stop 1
end if
    if ((nint(values(1)*10)) /= 5) then
    print *, "FAIL: want [5] got [", nint(values(1)*10), "]"
    stop 1
end if
    if ((nint(values(4)*10)) /= 20) then
    print *, "FAIL: want [20] got [", nint(values(4)*10), "]"
    stop 1
end if
end program array_implied_do_ordering_implied_fill_for_real
