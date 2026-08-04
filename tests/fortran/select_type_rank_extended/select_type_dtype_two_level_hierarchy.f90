! vybe-test: fortran/select_type_rank_extended/select_type_dtype_two_level_hierarchy
! origin: languages/fortran/tests/fortran/test_select_type_rank_extended.rs
program t
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 3 ]
type :: A
integer :: x = 1
end type A
type, extends(A) :: B
integer :: y = 2
end type B
type, extends(B) :: C
integer :: z = 3
end type C
class(A), allocatable :: obj
allocate(C :: obj)
select type(obj)
class is (C)
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 1) then
    print *, "FAIL: more than 1 line(s)"
    stop 1
end if
if ((obj%z) /= vybe_check_w(vybe_check_i)) then
    print *, "FAIL at ", vybe_check_i, " got [", obj%z, "]"
    stop 1
end if
class is (B)
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 1) then
    print *, "FAIL: more than 1 line(s)"
    stop 1
end if
if ((obj%y) /= vybe_check_w(vybe_check_i)) then
    print *, "FAIL at ", vybe_check_i, " got [", obj%y, "]"
    stop 1
end if
type is (A)
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 1) then
    print *, "FAIL: more than 1 line(s)"
    stop 1
end if
if ((obj%x) /= vybe_check_w(vybe_check_i)) then
    print *, "FAIL at ", vybe_check_i, " got [", obj%x, "]"
    stop 1
end if
end select
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program t
