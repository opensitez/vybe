! vybe-test: fortran/do_loop_extended/nested_do_product_table
! origin: languages/fortran/tests/fortran/test_do_loop_extended.rs
program t
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 46 ]
integer :: i, j, p
p = 1
do i = 1, 3
do j = 1, 3
p = p + i * j
end do
end do
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 1) then
    print *, "FAIL: more than 1 line(s)"
    stop 1
end if
if ((p) /= vybe_check_w(vybe_check_i)) then
    print *, "FAIL at ", vybe_check_i, " got [", p, "]"
    stop 1
end if
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program t
