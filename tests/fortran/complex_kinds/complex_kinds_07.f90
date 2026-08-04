! vybe-test: fortran/complex_kinds/complex_kinds_07
! origin: languages/fortran/tests/fortran/test_complex_kinds.rs
program p
complex :: a=(1.0,2.0)
print *, aimag(a)
end program p
