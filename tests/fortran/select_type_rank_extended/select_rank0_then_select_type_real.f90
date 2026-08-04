! vybe-test: fortran/select_type_rank_extended/select_rank0_then_select_type_real
! origin: languages/fortran/tests/fortran/test_select_type_rank_extended.rs
program t
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 4 ]
call inspect(4.5)
contains
subroutine inspect(x)
class(*), intent(in) :: x(..)
select rank(x)
rank(0)
select type(x)
type is (real)
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 1) then
    print *, "FAIL: more than 1 line(s)"
    stop 1
end if
if ((int(x)) /= vybe_check_w(vybe_check_i)) then
    print *, "FAIL at ", vybe_check_i, " got [", int(x), "]"
    stop 1
end if
class default
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 1) then
    print *, "FAIL: more than 1 line(s)"
    stop 1
end if
if ((0) /= vybe_check_w(vybe_check_i)) then
    print *, "FAIL at ", vybe_check_i, " got [", 0, "]"
    stop 1
end if
end select
rank default
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 1) then
    print *, "FAIL: more than 1 line(s)"
    stop 1
end if
if ((-1) /= vybe_check_w(vybe_check_i)) then
    print *, "FAIL at ", vybe_check_i, " got [", -1, "]"
    stop 1
end if
end select
end subroutine inspect
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program t
