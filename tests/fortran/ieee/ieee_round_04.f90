! vybe-test: fortran/ieee/ieee_round_04
! origin: languages/fortran/tests/fortran/test_ieee.rs
program p
use, intrinsic :: ieee_arithmetic
logical :: l
call ieee_get_rounding_mode(l)
end program p
