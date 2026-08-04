! vybe-test: fortran/ieee/ieee_rem_16
! origin: languages/fortran/tests/fortran/test_ieee.rs
program p
use, intrinsic :: ieee_arithmetic
print *, ieee_rem(7.0,3.0)
end program p
