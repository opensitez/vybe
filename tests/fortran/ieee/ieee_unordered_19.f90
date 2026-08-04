! vybe-test: fortran/ieee/ieee_unordered_19
! origin: languages/fortran/tests/fortran/test_ieee.rs
program p
use, intrinsic :: ieee_arithmetic
print *, ieee_unordered(1.0, 2.0)
end program p
