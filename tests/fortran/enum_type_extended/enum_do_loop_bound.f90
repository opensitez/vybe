! vybe-test: fortran/enum_type_extended/enum_do_loop_bound
! origin: languages/fortran/tests/fortran/test_enum_type_extended.rs
program t
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 15 ]
enum, bind(c)
enumerator :: START = 1, LAST = 5
end enum
integer :: i, s
s = 0
do i = START, LAST
s = s + i
end do
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 1) then
    print *, "FAIL: more than 1 line(s)"
    stop 1
end if
if ((s) /= vybe_check_w(vybe_check_i)) then
    print *, "FAIL at ", vybe_check_i, " got [", s, "]"
    stop 1
end if
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program t
