! vybe-test: fortran/subroutine_extended/pure_add_three_terms
! origin: languages/fortran/tests/fortran/test_subroutine_extended.rs
program t
if ((padd3(2, 3, 4)) /= 9) then
    print *, "FAIL: want [9] got [", padd3(2, 3, 4), "]"
    stop 1
end if
contains
pure function padd3(a, b, c) result(s)
integer, intent(in) :: a, b, c
integer :: s
s = a + b + c
end function padd3
end program t
