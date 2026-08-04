! vybe-test: fortran/array_constructor_repetition_counts/array_constructor_repetition_counts_singleton_implicit_default
! origin: languages/fortran/tests/fortran/test_array_constructor_repetition_counts.rs

program array_constructor_repetition_counts_singleton_implicit_default
    integer, allocatable :: values(:)
    values = (/ 6 * 1, 4, 5 * 2 /)
    if ((size(values)) /= 11) then
    print *, "FAIL: want [11] got [", size(values), "]"
    stop 1
end if
    if ((sum(values)) /= 24) then
    print *, "FAIL: want [24] got [", sum(values), "]"
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
