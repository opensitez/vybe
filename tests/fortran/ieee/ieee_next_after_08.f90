! vybe-test: fortran/ieee/ieee_next_after_08
! origin: languages/fortran/tests/fortran/test_ieee.rs
program p
use, intrinsic :: ieee_arithmetic
print *, ieee_next_after(1.0,2.0)
end program p
