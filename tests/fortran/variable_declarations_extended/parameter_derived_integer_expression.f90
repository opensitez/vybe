! vybe-test: fortran/variable_declarations_extended/parameter_derived_integer_expression
! origin: languages/fortran/tests/fortran/test_variable_declarations_extended.rs
program t
implicit none
integer, parameter :: base = 4
integer, parameter :: doubled = base * 2
if ((doubled) /= 8) then
    print *, "FAIL: want [8] got [", doubled, "]"
    stop 1
end if
end program t
