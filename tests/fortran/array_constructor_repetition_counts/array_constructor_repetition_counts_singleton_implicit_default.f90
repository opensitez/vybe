! vybe-test: fortran/array_constructor_repetition_counts/array_constructor_repetition_counts_singleton_implicit_default
! origin: languages/fortran/tests/fortran/test_array_constructor_repetition_counts.rs

program array_constructor_repetition_counts_singleton_implicit_default
    integer, allocatable :: values(:)
    values = (/ (1, i = 1, 6), 4, (2, i = 1, 5) /)
    if ((size(values)) /= 12) then
    print *, "FAIL: want [12] got [", size(values), "]"
    stop 1
end if
    if ((sum(values)) /= 20) then
    print *, "FAIL: want [20] got [", sum(values), "]"
    stop 1
end if
    if ((values(1)) /= 1) then
    print *, "FAIL: want [1] got [", values(1), "]"
    stop 1
end if
    if ((values(size(values))) /= 2) then
    print *, "FAIL: want [2] got [", values(size(values)), "]"
    stop 1
end if
end program array_constructor_repetition_counts_singleton_implicit_default
