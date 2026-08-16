! vybe-test: fortran/array_masked_array_operations/array_masked_array_operations_masked_redundant_mask
! origin: languages/fortran/tests/fortran/test_array_masked_array_operations.rs

program array_masked_array_operations_masked_redundant_mask
    integer :: values(4)
    integer :: result(4)
    values = (/ 2, 4, 6, 8 /)
    where (values > 3)
        result = values / 2
    else where
        result = values + 10
    end where
    if ((result(1)) /= 12) then
    print *, "FAIL: want [12] got [", result(1), "]"
    stop 1
end if
    if ((result(2)) /= 2) then
    print *, "FAIL: want [2] got [", result(2), "]"
    stop 1
end if
    if ((result(3)) /= 3) then
    print *, "FAIL: want [3] got [", result(3), "]"
    stop 1
end if
    if ((sum(result)) /= 21) then
    print *, "FAIL: want [21] got [", sum(result), "]"
    stop 1
end if
end program array_masked_array_operations_masked_redundant_mask
