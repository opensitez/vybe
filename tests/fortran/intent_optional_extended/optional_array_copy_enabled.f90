! vybe-test: fortran/intent_optional_extended/optional_array_copy_enabled
! origin: languages/fortran/tests/fortran/test_intent_optional_extended.rs
program t
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 9 ]
integer :: src(2), dst(2)
src = [4, 5]
call maybe_copy(src, dst, 2, .true.)
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 1) then
    print *, "FAIL: more than 1 line(s)"
    stop 1
end if
if ((sum(dst)) /= vybe_check_w(vybe_check_i)) then
    print *, "FAIL at ", vybe_check_i, " got [", sum(dst), "]"
    stop 1
end if
contains
subroutine maybe_copy(from, to, n, enable)
integer, intent(in) :: from(n), n
integer, intent(inout) :: to(n)
logical, intent(in), optional :: enable
integer :: i
if (present(enable) .and. enable) then
do i = 1, n
to(i) = from(i)
end do
else
to = 0
end if
end subroutine maybe_copy
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program t
