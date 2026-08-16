! vybe-test: fortran/array_constructor_repetition_counts/array_constructor_repetition_counts_even_odd_parity_mix
! origin: languages/fortran/tests/fortran/test_array_constructor_repetition_counts.rs

program array_constructor_repetition_counts_even_odd_parity_mix
    integer, allocatable :: values(:)
    values = (/ (2, i = 1, 3), (3, i = 1, 3), (2, i = 1, 2) /)
    if ((size(values)) /= 8) then
    print *, "FAIL: want [8] got [", size(values), "]"
    stop 1
end if
    if ((sum(values)) /= 19) then
    print *, "FAIL: want [19] got [", sum(values), "]"
    stop 1
end if
    if ((values(1)) /= 2) then
    print *, "FAIL: want [2] got [", values(1), "]"
    stop 1
end if
    if ((values(4)) /= 3) then
    print *, "FAIL: want [3] got [", values(4), "]"
    stop 1
end if
    if ((values(size(values))) /= 2) then
    print *, "FAIL: want [2] got [", values(size(values)), "]"
    stop 1
end if
end program array_constructor_repetition_counts_even_odd_parity_mix
