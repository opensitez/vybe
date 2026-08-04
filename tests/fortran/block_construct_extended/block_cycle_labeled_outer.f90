! vybe-test: fortran/block_construct_extended/block_cycle_labeled_outer
! origin: languages/fortran/tests/fortran/test_block_construct_extended.rs
program t
integer :: vybe_check_i = 0
integer :: vybe_check_w(4) = [ 11, 13, 31, 33 ]
integer :: i, j
outer: do i = 1, 3
do j = 1, 3
block
if (j == 2) cycle outer
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 4) then
    print *, "FAIL: more than 4 line(s)"
    stop 1
end if
if ((i * 10 + j) /= vybe_check_w(vybe_check_i)) then
    print *, "FAIL at ", vybe_check_i, " got [", i * 10 + j, "]"
    stop 1
end if
end block
end do
end do outer
if (vybe_check_i /= 4) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 4"
    stop 1
end if
end program t
