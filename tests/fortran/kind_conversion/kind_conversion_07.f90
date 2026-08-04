! vybe-test: fortran/kind_conversion/kind_conversion_07
! origin: languages/fortran/tests/fortran/test_kind_conversion.rs
program p
complex(kind=8) :: z
z = cmplx(1.0_8,2.0_8,kind=8)
print *, z
end program p
