! vybe-test: fortran/select_type_rank_extended/select_rank_module_procedure
! origin: languages/fortran/tests/fortran/test_select_type_rank_extended.rs
integer :: vybe_check_i = 0
integer :: vybe_check_w(2) = [ 3, 2 ]
module rankmod
contains
subroutine rows(x)
integer, intent(in) :: x(..)
select rank(x)
rank(2)
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 2) then
    print *, "FAIL: more than 2 line(s)"
    stop 1
end if
if ((size(x,1)) /= vybe_check_w(vybe_check_i)) then
    print *, "FAIL at ", vybe_check_i, " got [", size(x,1), "]"
    stop 1
end if
rank(1)
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 2) then
    print *, "FAIL: more than 2 line(s)"
    stop 1
end if
if ((size(x)) /= vybe_check_w(vybe_check_i)) then
    print *, "FAIL at ", vybe_check_i, " got [", size(x), "]"
    stop 1
end if
rank default
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 2) then
    print *, "FAIL: more than 2 line(s)"
    stop 1
end if
if ((0) /= vybe_check_w(vybe_check_i)) then
    print *, "FAIL at ", vybe_check_i, " got [", 0, "]"
    stop 1
end if
end select
end subroutine rows
end module rankmod
program t
use rankmod
call rows([1,2,3])
call rows(reshape([1,2,3,4,5,6],[2,3]))
if (vybe_check_i /= 2) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 2"
    stop 1
end if
end program t
