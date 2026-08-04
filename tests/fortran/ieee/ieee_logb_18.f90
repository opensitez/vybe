! vybe-test: fortran/ieee/ieee_logb_18
! origin: languages/fortran/tests/fortran/test_ieee.rs
program p
use, intrinsic :: ieee_arithmetic
print *, ieee_logb(8.0)
end program p
