! vybe-test: fortran/ieee/ieee_except_03
! origin: languages/fortran/tests/fortran/test_ieee.rs
program p
use, intrinsic :: ieee_exceptions
logical :: l
call ieee_get_halting_mode(ieee_divide_by_zero, l)
end program p
