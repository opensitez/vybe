! vybe-test: fortran/subroutine_extended/intent_in_dot_product_custom
! origin: languages/fortran/tests/fortran/test_subroutine_extended.rs
program t
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 11 ]
integer :: x(2), y(2)
x = [1, 2]
y = [3, 4]
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 1) then
    print *, "FAIL: more than 1 line(s)"
    stop 1
end if
if ((dot2(x, y, 2)) /= vybe_check_w(vybe_check_i)) then
    print *, "FAIL at ", vybe_check_i, " got [", dot2(x, y, 2), "]"
    stop 1
end if
contains
function dot2(u, v, n) result(s)
integer, intent(in) :: u(n), v(n), n
integer :: s, i
s = 0
do i = 1, n
s = s + u(i) * v(i)
end do
end function dot2
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program t
