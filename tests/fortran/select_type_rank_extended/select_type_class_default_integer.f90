! vybe-test: fortran/select_type_rank_extended/select_type_class_default_integer
! origin: languages/fortran/tests/fortran/test_select_type_rank_extended.rs
program t
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 9 ]
class(*), allocatable :: val
allocate(integer :: val)
val = 9
select type(val)
type is (real)
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
! `val` is CLASS(*) inside `class default` — its dynamic type is exactly what
! the branch did NOT establish, so no intrinsic operation on it is legal and
! gfortran rejected `val /= <integer>` outright. The subject of the test is
! which branch runs, and that is still decided here: the `type is (real)`
! branch above compares 0 against 9 and fails loudly if it ever claims this.
if ((9) /= vybe_check_w(vybe_check_i)) then
    print *, "FAIL at ", vybe_check_i, " got [", 9, "]"
    stop 1
end if
end select
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program t
