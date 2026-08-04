! vybe-test: fortran/complex_kinds/complex_kinds_06
! origin: languages/fortran/tests/fortran/test_complex_kinds.rs
program p
complex :: a=(1.0,2.0)
print *, conjg(a)
end program p
