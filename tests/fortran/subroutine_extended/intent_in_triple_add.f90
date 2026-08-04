! vybe-test: fortran/subroutine_extended/intent_in_triple_add
! origin: languages/fortran/tests/fortran/test_subroutine_extended.rs
program t
if ((add3(2, 3, 4)) /= 9) then
    print *, "FAIL: want [9] got [", add3(2, 3, 4), "]"
    stop 1
end if
contains
function add3(a, b, c) result(s)
integer, intent(in) :: a, b, c
integer :: s
s = a + b + c
end function add3
end program t
