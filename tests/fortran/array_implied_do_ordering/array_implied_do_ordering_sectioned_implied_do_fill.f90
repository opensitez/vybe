! vybe-test: fortran/array_implied_do_ordering/array_implied_do_ordering_sectioned_implied_do_fill
! origin: languages/fortran/tests/fortran/test_array_implied_do_ordering.rs

program array_implied_do_ordering_sectioned_implied_do_fill
    integer :: a(1:8)
    a = [(i, i = 1, 8)]
    a(3:6) = [(i*2, i = 1,4)]
    if ((a(1)) /= 1) then
    print *, "FAIL: want [1] got [", a(1), "]"
    stop 1
end if
    if ((a(5)) /= 8) then
    print *, "FAIL: want [8] got [", a(5), "]"
    stop 1
end if
    if ((sum(a)) /= 34) then
    print *, "FAIL: want [34] got [", sum(a), "]"
    stop 1
end if
end program array_implied_do_ordering_sectioned_implied_do_fill
