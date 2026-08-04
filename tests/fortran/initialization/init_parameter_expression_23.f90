! vybe-test: fortran/initialization/init_parameter_expression_23
! origin: languages/fortran/tests/fortran/test_initialization.rs
program p
integer, parameter :: one = 1
integer, parameter :: two = one + 1
integer, parameter :: three = two + one
print *, three
end program p
