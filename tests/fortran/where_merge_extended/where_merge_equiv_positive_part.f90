! vybe-test: fortran/where_merge_extended/where_merge_equiv_positive_part
! origin: languages/fortran/tests/fortran/test_where_merge_extended.rs
program t
integer :: vybe_check_i = 0
integer :: vybe_check_w(3) = [ 0, 5, 8 ]
integer :: a(4)=[-3,5,-1,8]
integer :: b(4)
integer :: i
do i=1,4
b(i)=merge(a(i),0,a(i)>0)
end do
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 3) then
    print *, "FAIL: more than 3 line(s)"
    stop 1
end if
if ((b(1)) /= vybe_check_w(vybe_check_i)) then
    print *, "FAIL at ", vybe_check_i, " got [", b(1), "]"
    stop 1
end if
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 3) then
    print *, "FAIL: more than 3 line(s)"
    stop 1
end if
if ((b(2)) /= vybe_check_w(vybe_check_i)) then
    print *, "FAIL at ", vybe_check_i, " got [", b(2), "]"
    stop 1
end if
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 3) then
    print *, "FAIL: more than 3 line(s)"
    stop 1
end if
if ((b(4)) /= vybe_check_w(vybe_check_i)) then
    print *, "FAIL at ", vybe_check_i, " got [", b(4), "]"
    stop 1
end if
if (vybe_check_i /= 3) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 3"
    stop 1
end if
end program t
