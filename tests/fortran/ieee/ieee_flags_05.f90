! vybe-test: fortran/ieee/ieee_flags_05
! origin: languages/fortran/tests/fortran/test_ieee.rs
program p
use, intrinsic :: ieee_exceptions
logical :: l
call ieee_get_flag(ieee_overflow, l)
end program p
