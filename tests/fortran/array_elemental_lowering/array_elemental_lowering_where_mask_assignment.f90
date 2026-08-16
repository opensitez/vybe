! vybe-test: fortran/array_elemental_lowering/array_elemental_lowering_where_mask_assignment
! origin: languages/fortran/tests/fortran/test_array_elemental_lowering.rs

program array_elemental_lowering_where_mask_assignment
    integer, allocatable :: values(:)
    integer, allocatable :: marked(:)
    values = (/ 1, 2, 3, 4, 5 /)
    marked = values
    where (values > 3)
        marked = 99
    end where
    if ((sum(marked)) /= 204) then
    print *, "FAIL: want [204] got [", sum(marked), "]"
    stop 1
end if
    if ((marked(3)) /= 3) then
    print *, "FAIL: want [3] got [", marked(3), "]"
    stop 1
end if
    if ((marked(5)) /= 99) then
    print *, "FAIL: want [99] got [", marked(5), "]"
    stop 1
end if
end program array_elemental_lowering_where_mask_assignment
