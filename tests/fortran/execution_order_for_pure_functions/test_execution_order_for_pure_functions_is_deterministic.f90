! vybe-test: fortran/execution_order_for_pure_functions/test_execution_order_for_pure_functions_is_deterministic
! origin: languages/fortran/tests/fortran/test_execution_order_for_pure_functions.rs

program test_execution_order_for_pure_functions
    integer :: value
    value = one_plus_two() + one_plus_two()
    if ((value) /= 6) then
    print *, "FAIL: want [6] got [", value, "]"
    stop 1
end if

contains
    pure function one_plus_two() result(r)
        integer :: r
        r = 1 + 2
    end function
end program test_execution_order_for_pure_functions
