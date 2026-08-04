! vybe-test: fortran/array_fill_pattern_internals/array_fill_pattern_internals_fill_strided_section
! origin: languages/fortran/tests/fortran/test_array_fill_pattern_internals.rs

program array_fill_pattern_internals_fill_strided_section
    integer :: values(1:7)
    values = 0
    values(1:7:2) = 3
    values(2:6:2) = 5
    if ((sum(values)) /= 27) then
    print *, "FAIL: want [27] got [", sum(values), "]"
    stop 1
end if
    if ((values(1)) /= 3) then
    print *, "FAIL: want [3] got [", values(1), "]"
    stop 1
end if
    if ((values(2)) /= 5) then
    print *, "FAIL: want [5] got [", values(2), "]"
    stop 1
end if
    if ((values(7)) /= 3) then
    print *, "FAIL: want [3] got [", values(7), "]"
    stop 1
end if
    if ((values(6)) /= 5) then
    print *, "FAIL: want [5] got [", values(6), "]"
    stop 1
end if
end program array_fill_pattern_internals_fill_strided_section
