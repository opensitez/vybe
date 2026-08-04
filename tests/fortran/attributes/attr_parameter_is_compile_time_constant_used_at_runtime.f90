! vybe-test: fortran/attributes/attr_parameter_is_compile_time_constant_used_at_runtime
! origin: languages/fortran/tests/fortran/test_attributes.rs

program attr_parameter_is_compile_time_constant_used_at_runtime
    integer, parameter :: n = 4
    integer :: a(n)
    a = [1, 2, 3, 4]
    if ((a(n)) /= 4) then
    print *, "FAIL: want [4] got [", a(n), "]"
    stop 1
end if
end program attr_parameter_is_compile_time_constant_used_at_runtime
