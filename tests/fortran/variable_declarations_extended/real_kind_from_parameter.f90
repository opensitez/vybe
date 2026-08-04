! vybe-test: fortran/variable_declarations_extended/real_kind_from_parameter
! origin: languages/fortran/tests/fortran/test_variable_declarations_extended.rs
program t
implicit none
integer, parameter :: rk = 8
real(kind=rk) :: y = 3.5_8
if (abs((y) - 3.5) > 1.0e-6) then
    print *, "FAIL: want [3.5] got [", y, "]"
    stop 1
end if
end program t
