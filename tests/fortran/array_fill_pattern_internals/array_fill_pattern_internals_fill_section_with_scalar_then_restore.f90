! vybe-test: fortran/array_fill_pattern_internals/array_fill_pattern_internals_fill_section_with_scalar_then_restore
! origin: languages/fortran/tests/fortran/test_array_fill_pattern_internals.rs

program t
    integer, allocatable :: values(:)
    values = (/ 10, 20, 30, 40, 50 /)
    values(2:4) = 0
    values(1) = values(5)
    if ((sum(values)) /= 100) then
    print *, "FAIL: want [100] got [", sum(values), "]"
    stop 1
end if
    if ((values(2)) /= 0) then
    print *, "FAIL: want [0] got [", values(2), "]"
    stop 1
end if
    if ((values(4)) /= 0) then
    print *, "FAIL: want [0] got [", values(4), "]"
    stop 1
end if
end program t
