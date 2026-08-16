! vybe-test: fortran/array_masked_array_operations/array_masked_array_operations_where_for_section_copy
! origin: languages/fortran/tests/fortran/test_array_masked_array_operations.rs

program array_masked_array_operations_where_for_section_copy
    integer :: source(1:6)
    integer :: result(1:6)
    source = (/ 9, 8, 7, 6, 5, 4 /)
    where (source > 6)
        result = source(1:6)
    elsewhere
        result = 0
    end where
    if ((sum(result)) /= 24) then
    print *, "FAIL: want [24] got [", sum(result), "]"
    stop 1
end if
    if ((result(1)) /= 9) then
    print *, "FAIL: want [9] got [", result(1), "]"
    stop 1
end if
    if ((result(4)) /= 0) then
    print *, "FAIL: want [0] got [", result(4), "]"
    stop 1
end if
    if ((result(6)) /= 0) then
    print *, "FAIL: want [0] got [", result(6), "]"
    stop 1
end if
end program array_masked_array_operations_where_for_section_copy
