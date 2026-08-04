! vybe-test: fortran/ieee/ieee_is_normal_12
! origin: languages/fortran/tests/fortran/test_ieee.rs
program p
use, intrinsic :: ieee_arithmetic
print *, ieee_is_normal(1.0)
end program p
