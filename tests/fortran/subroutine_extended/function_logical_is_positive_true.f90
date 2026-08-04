! vybe-test: fortran/subroutine_extended/function_logical_is_positive_true
! origin: languages/fortran/tests/fortran/test_subroutine_extended.rs
program t
if ((is_positive(7)) .neqv. .true.) then
    print *, "FAIL: want [true] got [", is_positive(7), "]"
    stop 1
end if
contains
logical function is_positive(n)
integer, intent(in) :: n
is_positive = n > 0
end function is_positive
end program t
