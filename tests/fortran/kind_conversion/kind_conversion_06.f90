! vybe-test: fortran/kind_conversion/kind_conversion_06
! origin: languages/fortran/tests/fortran/test_kind_conversion.rs
program p
real(kind=8) :: r
r = real(1, kind=8)
print *, r
end program p
