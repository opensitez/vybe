! vybe-test: fortran/real_ieee_components/modf_negative_truncates_toward_zero
! origin: languages/fortran/tests/fortran/test_real_ieee_components.rs
program t
real :: f
integer :: i
f = modf(-3.75, i)
print *, i, f
end program t
