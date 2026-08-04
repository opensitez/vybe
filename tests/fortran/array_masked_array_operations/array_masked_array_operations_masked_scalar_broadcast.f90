! vybe-test: fortran/array_masked_array_operations/array_masked_array_operations_masked_scalar_broadcast
! origin: languages/fortran/tests/fortran/test_array_masked_array_operations.rs

program array_masked_array_operations_masked_scalar_broadcast
    integer :: values(3)
    integer :: result(3)
    values = (/ 1, 2, 3 /)
    where (values /= 0)
        result = 7
    end where
    if ((sum(result)) /= 21) then
    print *, "FAIL: want [21] got [", sum(result), "]"
    stop 1
end if
    if ((result(1)) /= 7) then
    print *, "FAIL: want [7] got [", result(1), "]"
    stop 1
end if
    if ((result(3)) /= 7) then
    print *, "FAIL: want [7] got [", result(3), "]"
    stop 1
end if
end program array_masked_array_operations_masked_scalar_broadcast
