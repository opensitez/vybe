! vybe-test: fortran/subroutine_extended/intent_in_product_pair
! origin: languages/fortran/tests/fortran/test_subroutine_extended.rs
program t
if ((mul2(6, 7)) /= 42) then
    print *, "FAIL: want [42] got [", mul2(6, 7), "]"
    stop 1
end if
contains
function mul2(x, y) result(p)
integer, intent(in) :: x, y
integer :: p
p = x * y
end function mul2
end program t
