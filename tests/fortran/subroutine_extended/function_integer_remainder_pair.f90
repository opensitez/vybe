! vybe-test: fortran/subroutine_extended/function_integer_remainder_pair
! origin: languages/fortran/tests/fortran/test_subroutine_extended.rs
program t
if ((rem_pair(17, 5)) /= 2) then
    print *, "FAIL: want [2] got [", rem_pair(17, 5), "]"
    stop 1
end if
contains
integer function rem_pair(a, b)
integer, intent(in) :: a, b
rem_pair = mod(a, b)
end function rem_pair
end program t
