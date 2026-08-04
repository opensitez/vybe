! vybe-test: fortran/array_masked_array_operations/array_masked_array_operations_mask_for_logical_condition
! origin: languages/fortran/tests/fortran/test_array_masked_array_operations.rs

program array_masked_array_operations_mask_for_logical_condition
    integer :: values(7)
    integer :: selected(7)
    integer :: hits
    values = (/ 2, 4, 6, 8, 10, 12, 14 /)
    selected = 0
    where (mod(values, 4) == 0)
        selected = 1
    end where
    hits = sum(selected)
    if ((hits) /= 3) then
    print *, "FAIL: want [3] got [", hits, "]"
    stop 1
end if
    if ((count(selected == 1)) /= 3) then
    print *, "FAIL: want [3] got [", count(selected == 1), "]"
    stop 1
end if
    if ((selected(2)) /= 1) then
    print *, "FAIL: want [1] got [", selected(2), "]"
    stop 1
end if
    if ((selected(3)) /= 0) then
    print *, "FAIL: want [0] got [", selected(3), "]"
    stop 1
end if
end program array_masked_array_operations_mask_for_logical_condition
