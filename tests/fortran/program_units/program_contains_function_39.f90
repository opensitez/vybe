! vybe-test: fortran/program_units/program_contains_function_39
! origin: languages/fortran/tests/fortran/test_program_units.rs
program p
print *, f()
contains
integer function f()
f = 1
end function f
end program p
