! vybe-test: fortran/kind_conversion/kind_conversion_09
! origin: languages/fortran/tests/fortran/test_kind_conversion.rs
program p
real :: r
r = transfer(1, r)
print *, r
end program p
