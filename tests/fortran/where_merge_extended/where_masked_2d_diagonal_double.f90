! vybe-test: fortran/where_merge_extended/where_masked_2d_diagonal_double
! origin: languages/fortran/tests/fortran/test_where_merge_extended.rs
program t
integer :: vybe_check_i = 0
integer :: vybe_check_w(3) = [ 10, 20, 0 ]
integer :: m(3,3)
integer :: i,j
m=0
do i=1,3
m(i,i)=i
end do
where (m>0)
m=m*10
end where
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 3) then
    print *, "FAIL: more than 3 line(s)"
    stop 1
end if
if ((m(1,1)) /= vybe_check_w(vybe_check_i)) then
    print *, "FAIL at ", vybe_check_i, " got [", m(1,1), "]"
    stop 1
end if
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 3) then
    print *, "FAIL: more than 3 line(s)"
    stop 1
end if
if ((m(2,2)) /= vybe_check_w(vybe_check_i)) then
    print *, "FAIL at ", vybe_check_i, " got [", m(2,2), "]"
    stop 1
end if
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 3) then
    print *, "FAIL: more than 3 line(s)"
    stop 1
end if
if ((m(1,2)) /= vybe_check_w(vybe_check_i)) then
    print *, "FAIL at ", vybe_check_i, " got [", m(1,2), "]"
    stop 1
end if
if (vybe_check_i /= 3) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 3"
    stop 1
end if
end program t
