! vybe-test: fortran/intent_optional_extended/intent_inout_optional_scale_array
! origin: languages/fortran/tests/fortran/test_intent_optional_extended.rs
program t
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 14 ]
integer :: a(2)
a = [3, 4]
call scale_if(a, 2, 2)
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 1) then
    print *, "FAIL: more than 1 line(s)"
    stop 1
end if
if ((sum(a)) /= vybe_check_w(vybe_check_i)) then
    print *, "FAIL at ", vybe_check_i, " got [", sum(a), "]"
    stop 1
end if
contains
subroutine scale_if(v, n, k)
integer, intent(inout) :: v(n)
integer, intent(in) :: n
integer, intent(in), optional :: k
integer :: i
if (present(k)) then
do i = 1, n
v(i) = v(i) * k
end do
end if
end subroutine scale_if
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program t
