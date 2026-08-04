! vybe-test: fortran/ieee/ieee_support_nan_15
! origin: languages/fortran/tests/fortran/test_ieee.rs
program p
use, intrinsic :: ieee_arithmetic
print *, ieee_support_nan(1.0)
end program p
