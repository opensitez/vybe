! vybe-test: fortran/ieee/ieee_datatype_20
! origin: languages/fortran/tests/fortran/test_ieee.rs
program p
use, intrinsic :: ieee_arithmetic
print *, ieee_support_datatype((1.0,2.0))
end program p
