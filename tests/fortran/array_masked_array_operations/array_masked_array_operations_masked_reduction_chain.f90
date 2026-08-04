! vybe-test: fortran/array_masked_array_operations/array_masked_array_operations_masked_reduction_chain
! origin: languages/fortran/tests/fortran/test_array_masked_array_operations.rs

program array_masked_array_operations_masked_reduction_chain
    integer :: values(4)
    integer :: result(4)
    values = (/ 5, 10, 15, 20 /)
    where (values >= 10)
        result = values / 5
    end where
    if ((sum(result)) /= 9) then
    print *, "FAIL: want [9] got [", sum(result), "]"
    stop 1
end if
    if ((count(result == 0)) /= 2) then
    print *, "FAIL: want [2] got [", count(result == 0), "]"
    stop 1
end if
    if ((sum(merge(1, 0, result /= 0))) /= 3) then
    print *, "FAIL: want [3] got [", sum(merge(1, 0, result /= 0)), "]"
    stop 1
end if
end program array_masked_array_operations_masked_reduction_chain
