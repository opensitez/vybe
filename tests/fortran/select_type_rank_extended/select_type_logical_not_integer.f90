! vybe-test: fortran/select_type_rank_extended/select_type_logical_not_integer
! origin: languages/fortran/tests/fortran/test_select_type_rank_extended.rs
program t
integer :: vybe_check_i = 0
logical :: vybe_check_w(1) = [ .true. ]
class(*), allocatable :: val
allocate(logical :: val)
val = .true.
select type(val)
type is (integer)
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 1) then
    print *, "FAIL: more than 1 line(s)"
    stop 1
end if
if ((0) .neqv. vybe_check_w(vybe_check_i)) then
    print *, "FAIL at ", vybe_check_i, " got [", 0, "]"
    stop 1
end if
type is (logical)
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 1) then
    print *, "FAIL: more than 1 line(s)"
    stop 1
end if
if ((val) .neqv. vybe_check_w(vybe_check_i)) then
    print *, "FAIL at ", vybe_check_i, " got [", val, "]"
    stop 1
end if
end select
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program t
