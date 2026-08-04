! vybe-test: fortran/kind_conversion/kind_conversion_03
! origin: languages/fortran/tests/fortran/test_kind_conversion.rs
program p
double precision :: d
d = dble(1.0)
print *, d
end program p
