! vybe-test: fortran/real_ieee_components/modf_zero_yields_zero_fraction
! origin: languages/fortran/tests/fortran/test_real_ieee_components.rs
program t
real :: f
integer :: i
f = modf(0.0, i)
print *, i, f
end program t
