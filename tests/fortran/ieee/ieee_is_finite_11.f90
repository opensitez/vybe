! vybe-test: fortran/ieee/ieee_is_finite_11
! origin: languages/fortran/tests/fortran/test_ieee.rs
program p
use, intrinsic :: ieee_arithmetic
print *, ieee_is_finite(1.0)
end program p
