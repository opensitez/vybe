! vybe-test: fortran/array_elemental_lowering/array_elemental_lowering_cascade_reassignments
! origin: languages/fortran/tests/fortran/test_array_elemental_lowering.rs

program array_elemental_lowering_cascade_reassignments
    integer, allocatable :: values(:)
    integer :: result
    values = (/ 1, 2, 3, 4 /)
    values = values + 1
    values = values * 2
    result = sum(values)
    if ((result) /= 28) then
    print *, "FAIL: want [28] got [", result, "]"
    stop 1
end if
    if ((values(1)) /= 4) then
    print *, "FAIL: want [4] got [", values(1), "]"
    stop 1
end if
    if ((values(4)) /= 10) then
    print *, "FAIL: want [10] got [", values(4), "]"
    stop 1
end if
end program array_elemental_lowering_cascade_reassignments
