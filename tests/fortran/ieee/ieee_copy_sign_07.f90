! vybe-test: fortran/ieee/ieee_copy_sign_07
! origin: languages/fortran/tests/fortran/test_ieee.rs
program p
use, intrinsic :: ieee_arithmetic
print *, ieee_copy_sign(1.0,-2.0)
end program p
