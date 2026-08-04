! vybe-test: fortran/program_units/program_statement_fn_08
! origin: languages/fortran/tests/fortran/test_program_units.rs
program p
implicit none
integer :: f
f(x) = x + 1
print *, f(1)
end program p
