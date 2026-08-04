! vybe-test: fortran/block_construct_extended/block_exit_after_three_prints
! origin: languages/fortran/tests/fortran/test_block_construct_extended.rs
program t
integer :: vybe_check_i = 0
integer :: vybe_check_w(3) = [ 1, 2, 3 ]
integer :: k
outer: do k = 1, 20
block
if (k > 3) exit outer
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 3) then
    print *, "FAIL: more than 3 line(s)"
    stop 1
end if
if ((k) /= vybe_check_w(vybe_check_i)) then
    print *, "FAIL at ", vybe_check_i, " got [", k, "]"
    stop 1
end if
end block
end do outer
if (vybe_check_i /= 3) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 3"
    stop 1
end if
end program t
