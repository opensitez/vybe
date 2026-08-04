! vybe-test: fortran/ieee/ieee_rint_17
! origin: languages/fortran/tests/fortran/test_ieee.rs
program p
use, intrinsic :: ieee_arithmetic
print *, ieee_rint(1.6)
end program p
