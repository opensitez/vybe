! vybe-test: fortran/arrays/read_only_array_param_function_works_in_expression_runtime
! origin: languages/fortran/tests/fortran/test_arrays.rs

program test
    real :: data(3) = [1.1, 1.5, 3.0]
    if (abs((first_value(data) + 1.0) - 2.1) > 1.0e-6) then
    print *, "FAIL: want [2.1] got [", first_value(data) + 1.0, "]"
    stop 1
end if
contains
    real function first_value(values) result(out)
        real, intent(in) :: values(:)
        real :: out
        out = values(1)
    end function first_value
end program test
