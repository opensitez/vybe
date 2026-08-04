! vybe-test: fortran/complex_kinds/complex_kinds_04
! origin: languages/fortran/tests/fortran/test_complex_kinds.rs
program p
complex(kind=8) :: z=(1.0_8,2.0_8)
print *, z
end program p
