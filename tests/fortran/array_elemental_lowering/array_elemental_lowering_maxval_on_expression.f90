! vybe-test: fortran/array_elemental_lowering/array_elemental_lowering_maxval_on_expression
! origin: languages/fortran/tests/fortran/test_array_elemental_lowering.rs

program array_elemental_lowering_maxval_on_expression
    integer, allocatable :: values(:)
    values = (/ -9, 12, 4, 18, 3 /)
    if ((maxval(values)) /= 18) then
    print *, "FAIL: want [18] got [", maxval(values), "]"
    stop 1
end if
    if ((maxval(abs(values))) /= 18) then
    print *, "FAIL: want [18] got [", maxval(abs(values)), "]"
    stop 1
end if
    if ((size(pack(values, values > 5))) /= 2) then
    print *, "FAIL: want [2] got [", size(pack(values, values > 5)), "]"
    stop 1
end if
end program array_elemental_lowering_maxval_on_expression
