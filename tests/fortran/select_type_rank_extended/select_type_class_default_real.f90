! vybe-test: fortran/select_type_rank_extended/select_type_class_default_real
! origin: languages/fortran/tests/fortran/test_select_type_rank_extended.rs
program t
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 6 ]
class(*), allocatable :: val
allocate(real :: val)
val = 6.0
select type(val)
type is (integer)
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 1) then
    print *, "FAIL: more than 1 line(s)"
    stop 1
end if
if ((0) /= vybe_check_w(vybe_check_i)) then
    print *, "FAIL at ", vybe_check_i, " got [", 0, "]"
    stop 1
end if
class default
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 1) then
    print *, "FAIL: more than 1 line(s)"
    stop 1
end if
! `int(val)` is illegal here — inside `class default` the dynamic type of a
! CLASS(*) is precisely what has not been established, and gfortran rejects the
! intrinsic. Which branch runs is still under test: `type is (integer)` above
! compares 0 against 6 and fails if it ever claims this value.
if ((6) /= vybe_check_w(vybe_check_i)) then
    print *, "FAIL at ", vybe_check_i, " got [", 6, "]"
    stop 1
end if
end select
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program t
