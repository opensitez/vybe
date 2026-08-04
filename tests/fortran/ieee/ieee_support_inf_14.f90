! vybe-test: fortran/ieee/ieee_support_inf_14
! origin: languages/fortran/tests/fortran/test_ieee.rs
program p
use, intrinsic :: ieee_arithmetic
print *, ieee_support_inf(1.0)
end program p
