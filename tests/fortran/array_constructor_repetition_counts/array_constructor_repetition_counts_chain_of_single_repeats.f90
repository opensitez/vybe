! vybe-test: fortran/array_constructor_repetition_counts/array_constructor_repetition_counts_chain_of_single_repeats
! origin: languages/fortran/tests/fortran/test_array_constructor_repetition_counts.rs

program array_constructor_repetition_counts_chain_of_single_repeats
    integer, allocatable :: values(:)
    values = (/ (8, i = 1, 1), (1, i = 1, 1), (6, i = 1, 1), (4, i = 1, 1) /)
    if ((size(values)) /= 4) then
    print *, "FAIL: want [4] got [", size(values), "]"
    stop 1
end if
    if ((sum(values)) /= 19) then
    print *, "FAIL: want [19] got [", sum(values), "]"
    stop 1
end if
    if ((values(1)) /= 8) then
    print *, "FAIL: want [8] got [", values(1), "]"
    stop 1
end if
    if ((values(size(values))) /= 4) then
    print *, "FAIL: want [4] got [", values(size(values)), "]"
    stop 1
end if
end program array_constructor_repetition_counts_chain_of_single_repeats
