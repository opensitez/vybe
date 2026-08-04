! vybe-test: fortran/array_elemental_lowering/array_elemental_lowering_minval_and_index
! origin: languages/fortran/tests/fortran/test_array_elemental_lowering.rs

program array_elemental_lowering_minval_and_index
    integer, allocatable :: values(:)
    values = (/ 7, 3, 9, 3, 1 /)
    if ((minval(values)) /= 1) then
    print *, "FAIL: want [1] got [", minval(values), "]"
    stop 1
end if
    if ((minval(values) + maxval(values)) /= 10) then
    print *, "FAIL: want [10] got [", minval(values) + maxval(values), "]"
    stop 1
end if
    if ((maxloc(values)) /= 5) then
    print *, "FAIL: want [5] got [", maxloc(values), "]"
    stop 1
end if
end program array_elemental_lowering_minval_and_index
