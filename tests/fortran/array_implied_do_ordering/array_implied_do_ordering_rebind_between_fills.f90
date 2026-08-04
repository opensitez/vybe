! vybe-test: fortran/array_implied_do_ordering/array_implied_do_ordering_rebind_between_fills
! origin: languages/fortran/tests/fortran/test_array_implied_do_ordering.rs

program array_implied_do_ordering_rebind_between_fills
    integer :: a(4)
    integer :: b(4)
    a = [(i, i = 1,4)]
    b = [(i*2, i = 1,4)]
    if ((sum(a)) /= 10) then
    print *, "FAIL: want [10] got [", sum(a), "]"
    stop 1
end if
    if ((sum(b)) /= 20) then
    print *, "FAIL: want [20] got [", sum(b), "]"
    stop 1
end if
    if ((b(2) - a(2)) /= 2) then
    print *, "FAIL: want [2] got [", b(2) - a(2), "]"
    stop 1
end if
end program array_implied_do_ordering_rebind_between_fills
