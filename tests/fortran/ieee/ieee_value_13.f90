! vybe-test: fortran/ieee/ieee_value_13
! origin: languages/fortran/tests/fortran/test_ieee.rs
program p
use, intrinsic :: ieee_arithmetic
print *, ieee_value(1.0, ieee_positive_inf)
end program p
