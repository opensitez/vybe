! vybe-test: fortran/select_type_rank_extended/select_type_complex_not_matched_default
! origin: languages/fortran/tests/fortran/test_select_type_rank_extended.rs
program t
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 3 ]
class(*), allocatable :: val
allocate(complex :: val)
val = (1.0, 2.0)
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
! `aimag(val)` is illegal on a CLASS(*) — inside `class default` the dynamic
! type is the one thing the branch has NOT established, and gfortran rejects
! it. The subject survives: a complex must reach `class default`, and both the
! `integer` and `real` branches above compare 0 against 3 and fail if either
! ever claims it.
if ((3) /= vybe_check_w(vybe_check_i)) then
    print *, "FAIL at ", vybe_check_i, " got [", 3, "]"
    stop 1
end if
end select
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program t
