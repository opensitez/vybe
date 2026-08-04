! vybe-test: fortran/kind_conversion/kind_conversion_04
! origin: languages/fortran/tests/fortran/test_kind_conversion.rs
program p
complex :: z
z = cmplx(1.0,2.0)
print *, z
end program p
