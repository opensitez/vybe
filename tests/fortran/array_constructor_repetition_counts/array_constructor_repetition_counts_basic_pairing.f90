! vybe-test: fortran/array_constructor_repetition_counts/array_constructor_repetition_counts_basic_pairing
! origin: languages/fortran/tests/fortran/test_array_constructor_repetition_counts.rs

program array_constructor_repetition_counts_basic_pairing
    integer, allocatable :: values(:)
    values = (/ (10, i = 1, 2), (20, i = 1, 3) /)
    if ((size(values)) /= 5) then
    print *, "FAIL: want [5] got [", size(values), "]"
    stop 1
end if
    if ((sum(values)) /= 80) then
    print *, "FAIL: want [80] got [", sum(values), "]"
    stop 1
end if
    if ((values(1)) /= 10) then
    print *, "FAIL: want [10] got [", values(1), "]"
    stop 1
end if
    if ((values(size(values))) /= 20) then
    print *, "FAIL: want [20] got [", values(size(values)), "]"
    stop 1
end if
end program array_constructor_repetition_counts_basic_pairing
