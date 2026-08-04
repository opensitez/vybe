! vybe-test: fortran/array_constructor_repetition_counts/array_constructor_repetition_counts_nested_constructor_roundtrip
! origin: languages/fortran/tests/fortran/test_array_constructor_repetition_counts.rs

program array_constructor_repetition_counts_nested_constructor_roundtrip
    integer, allocatable :: values(:)
    values = (/ 2 * (3 * 1), 1 * (2 * 2), 1 * (2 + 1) /)
    if ((size(values)) /= 4) then
    print *, "FAIL: want [4] got [", size(values), "]"
    stop 1
end if
    if ((sum(values)) /= 9) then
    print *, "FAIL: want [9] got [", sum(values), "]"
    stop 1
end if
    if ((values(2)) /= 3) then
    print *, "FAIL: want [3] got [", values(2), "]"
    stop 1
end if
    if ((values(size(values))) /= 3) then
    print *, "FAIL: want [3] got [", values(size(values)), "]"
    stop 1
end if
end program array_constructor_repetition_counts_nested_constructor_roundtrip
