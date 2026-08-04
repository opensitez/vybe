! vybe-test: fortran/variable_declarations_extended/save_initialized_scalar
! origin: languages/fortran/tests/fortran/test_variable_declarations_extended.rs
program t
implicit none
integer, save :: counter = 99
if ((counter) /= 99) then
    print *, "FAIL: want [99] got [", counter, "]"
    stop 1
end if
end program t
