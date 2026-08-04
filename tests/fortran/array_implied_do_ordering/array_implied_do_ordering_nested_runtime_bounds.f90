! vybe-test: fortran/array_implied_do_ordering/array_implied_do_ordering_nested_runtime_bounds
! origin: languages/fortran/tests/fortran/test_array_implied_do_ordering.rs

program array_implied_do_ordering_nested_runtime_bounds
    integer, allocatable :: values(:)
    integer :: i_start
    integer :: j_end
    i_start = 1
    j_end = 4
    values = (/ (i + j, i = i_start, 2, j = 1, j_end) /)
    if ((size(values)) /= 8) then
    print *, "FAIL: want [8] got [", size(values), "]"
    stop 1
end if
    if ((sum(values)) /= 32) then
    print *, "FAIL: want [32] got [", sum(values), "]"
    stop 1
end if
    if ((values(1)) /= 2) then
    print *, "FAIL: want [2] got [", values(1), "]"
    stop 1
end if
    if ((values(size(values))) /= 6) then
    print *, "FAIL: want [6] got [", values(size(values)), "]"
    stop 1
end if
    if ((values(5)) /= 3) then
    print *, "FAIL: want [3] got [", values(5), "]"
    stop 1
end if
end program array_implied_do_ordering_nested_runtime_bounds
