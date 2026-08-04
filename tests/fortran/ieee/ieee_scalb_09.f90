! vybe-test: fortran/ieee/ieee_scalb_09
! origin: languages/fortran/tests/fortran/test_ieee.rs
program p
use, intrinsic :: ieee_arithmetic
print *, ieee_scalb(1.0,2)
end program p
