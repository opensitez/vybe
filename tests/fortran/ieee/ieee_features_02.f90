! vybe-test: fortran/ieee/ieee_features_02
! origin: languages/fortran/tests/fortran/test_ieee.rs
program p
use, intrinsic :: ieee_features
print *, ieee_support_underflow_control(1.0)
end program p
