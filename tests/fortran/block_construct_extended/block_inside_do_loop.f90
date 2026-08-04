! vybe-test: fortran/block_construct_extended/block_inside_do_loop
! origin: languages/fortran/tests/fortran/test_block_construct_extended.rs
program t
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 14 ]
integer :: i, s
s = 0
do i = 1, 3
block
integer :: sq
sq = i * i
s = s + sq
end block
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
