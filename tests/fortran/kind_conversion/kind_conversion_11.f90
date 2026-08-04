! vybe-test: fortran/kind_conversion/kind_conversion_11
! origin: languages/fortran/tests/fortran/test_kind_conversion.rs
program p
integer*4 :: i
real*8 :: r
i = int(1.5, 4)
r = real(i, 8)
print *, i
print *, r
end program p
