! vybe-test: fortran/kind_conversion/kind_conversion_12
! origin: languages/fortran/tests/fortran/test_kind_conversion.rs
program p
real, dimension(3) :: r
integer, dimension(3) :: i
i = int(r)
r = real(i)
print *, i(1)
print *, i(2)
print *, i(3)
print *, r(1)
end program p
