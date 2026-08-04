! vybe-test: fortran/ieee/ieee_classify_06
! origin: languages/fortran/tests/fortran/test_ieee.rs
program p
use, intrinsic :: ieee_arithmetic
print *, ieee_class(1.0)
end program p
