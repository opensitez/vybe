! vybe-test: fortran/ieee/ieee_arith_01
! origin: languages/fortran/tests/fortran/test_ieee.rs
program p
use, intrinsic :: ieee_arithmetic
print *, ieee_support_datatype(1.0)
end program p
