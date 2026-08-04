! vybe-test: fortran/array_elemental_lowering/array_elemental_lowering_count_with_nested_mask
! origin: languages/fortran/tests/fortran/test_array_elemental_lowering.rs

program array_elemental_lowering_count_with_nested_mask
    integer, allocatable :: values(:)
    values = (/ 0, 1, 2, 3, -1, -2 /)
    if ((count(values >= 0)) /= 4) then
    print *, "FAIL: want [4] got [", count(values >= 0), "]"
    stop 1
end if
    if ((count((values == 0) .or. (values == -1))) /= 2) then
    print *, "FAIL: want [2] got [", count((values == 0) .or. (values == -1)), "]"
    stop 1
end if
    if ((maxval(abs(values))) /= 3) then
    print *, "FAIL: want [3] got [", maxval(abs(values)), "]"
    stop 1
end if
end program array_elemental_lowering_count_with_nested_mask
