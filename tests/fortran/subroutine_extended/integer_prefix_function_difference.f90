! vybe-test: fortran/subroutine_extended/integer_prefix_function_difference
! origin: languages/fortran/tests/fortran/test_subroutine_extended.rs
program t
if ((diff(15, 6)) /= 9) then
    print *, "FAIL: want [9] got [", diff(15, 6), "]"
    stop 1
end if
contains
integer function diff(a, b)
integer, intent(in) :: a, b
diff = a - b
end function diff
end program t
