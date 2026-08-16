! vybe-test: fortran/intent_optional_extended/intent_in_rank_one_stride_sum
! origin: languages/fortran/tests/fortran/test_intent_optional_extended.rs
program t
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 9 ]
integer :: v(5)
v = [1, 2, 3, 4, 5]
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 1) then
    print *, "FAIL: more than 1 line(s)"
    stop 1
end if
if ((stride_sum(v, 5, 2)) /= vybe_check_w(vybe_check_i)) then
    print *, "FAIL at ", vybe_check_i, " got [", stride_sum(v, 5, 2), "]"
    stop 1
end if
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
contains
function stride_sum(a, n, step) result(s)
integer, intent(in) :: a(n), n, step
integer :: s, i
s = 0
do i = 1, n, step
s = s + a(i)
end do
end function stride_sum
end program t
