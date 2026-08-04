! vybe-test: fortran/array_implied_do_ordering/array_implied_do_ordering_runtime_empty_sequence
! origin: languages/fortran/tests/fortran/test_array_implied_do_ordering.rs

program array_implied_do_ordering_runtime_empty_sequence
    integer, allocatable :: values(:)
    values = (/ (i, i = 3, 1, 2) /)
    if ((size(values)) /= 0) then
    print *, "FAIL: want [0] got [", size(values), "]"
    stop 1
end if
    if ((merge(1, 0, size(values) == 0)) /= 1) then
    print *, "FAIL: want [1] got [", merge(1, 0, size(values) == 0), "]"
    stop 1
end if
    if ((sum(values)) /= 0) then
    print *, "FAIL: want [0] got [", sum(values), "]"
    stop 1
end if
end program array_implied_do_ordering_runtime_empty_sequence
