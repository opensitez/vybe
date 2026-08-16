! vybe-test: fortran/array_elemental_lowering/array_elemental_lowering_overlapping_section_with_rhs_temp_semantics
! origin: languages/fortran/tests/fortran/test_array_elemental_lowering.rs

program t
    integer, allocatable :: values(:)
    values = (/ 1, 2, 3, 4, 5, 6 /)
    values(2:5) = values(1:4) + values(2:5)
    if ((size(values)) /= 6) then
    print *, "FAIL: want [6] got [", size(values), "]"
    stop 1
end if
    if ((sum(values)) /= 31) then
    print *, "FAIL: want [31] got [", sum(values), "]"
    stop 1
end if
    if ((values(1)) /= 1) then
    print *, "FAIL: want [1] got [", values(1), "]"
    stop 1
end if
    if ((values(2)) /= 3) then
    print *, "FAIL: want [3] got [", values(2), "]"
    stop 1
end if
    if ((values(5)) /= 9) then
    print *, "FAIL: want [9] got [", values(5), "]"
    stop 1
end if
end program t
