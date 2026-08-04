! vybe-test: fortran/block_construct_extended/block_nested_with_exit_inner_only
! origin: languages/fortran/tests/fortran/test_block_construct_extended.rs
program t
integer :: vybe_check_i = 0
integer :: vybe_check_w(5) = [ 11, 21, 31, 41, 51 ]
integer :: i
outer: do i = 1, 5
block
integer :: j
do j = 1, 3
if (j == 2) exit
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 5) then
    print *, "FAIL: more than 5 line(s)"
    stop 1
end if
if ((i * 10 + j) /= vybe_check_w(vybe_check_i)) then
    print *, "FAIL at ", vybe_check_i, " got [", i * 10 + j, "]"
    stop 1
end if
end do
end block
end do outer
if (vybe_check_i /= 5) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 5"
    stop 1
end if
end program t
