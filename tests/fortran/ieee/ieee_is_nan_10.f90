! vybe-test: fortran/ieee/ieee_is_nan_10
! origin: languages/fortran/tests/fortran/test_ieee.rs
program p
use, intrinsic :: ieee_arithmetic
real :: x
x = 0.0/0.0
print *, ieee_is_nan(x)
end program p
