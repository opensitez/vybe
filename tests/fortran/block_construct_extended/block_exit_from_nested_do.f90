! vybe-test: fortran/block_construct_extended/block_exit_from_nested_do
! origin: languages/fortran/tests/fortran/test_block_construct_extended.rs
program t
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 11 ]
integer :: i, j
outer: do i = 1, 5
do j = 1, 5
block
if (j == 2) exit outer
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 1) then
    print *, "FAIL: more than 1 line(s)"
    stop 1
end if
if ((i * 10 + j) /= vybe_check_w(vybe_check_i)) then
    print *, "FAIL at ", vybe_check_i, " got [", i * 10 + j, "]"
    stop 1
end if
end block
end do
end do outer
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program t
