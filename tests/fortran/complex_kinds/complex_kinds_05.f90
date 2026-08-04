! vybe-test: fortran/complex_kinds/complex_kinds_05
! origin: languages/fortran/tests/fortran/test_complex_kinds.rs
program p
complex :: a=(1.0,2.0), b=(3.0,4.0)
print *, a+b
end program p
