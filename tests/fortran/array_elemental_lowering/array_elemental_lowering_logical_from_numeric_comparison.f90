! vybe-test: fortran/array_elemental_lowering/array_elemental_lowering_logical_from_numeric_comparison
! origin: languages/fortran/tests/fortran/test_array_elemental_lowering.rs

program array_elemental_lowering_logical_from_numeric_comparison
    integer, allocatable :: values(:)
    integer :: positives
    values = (/ -3, -1, 0, 4, 2 /)
    positives = count(values > 0)
    if ((positives) /= 2) then
    print *, "FAIL: want [2] got [", positives, "]"
    stop 1
end if
    if ((maxval(merge(1, 0, values > 0))) /= 1) then
    print *, "FAIL: want [1] got [", maxval(merge(1, 0, values > 0)), "]"
    stop 1
end if
end program array_elemental_lowering_logical_from_numeric_comparison
