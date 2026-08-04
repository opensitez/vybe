! vybe-test: fortran/variable_declarations_extended/logical_parameter_constant
! origin: languages/fortran/tests/fortran/test_variable_declarations_extended.rs
program t
implicit none
logical, parameter :: ok = .true.
if ((ok) .neqv. .true.) then
    print *, "FAIL: want [true] got [", ok, "]"
    stop 1
end if
end program t
