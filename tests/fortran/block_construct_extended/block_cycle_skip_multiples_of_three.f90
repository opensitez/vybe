! vybe-test: fortran/block_construct_extended/block_cycle_skip_multiples_of_three
! origin: languages/fortran/tests/fortran/test_block_construct_extended.rs
program t
integer :: vybe_check_i = 0
integer :: vybe_check_w(6) = [ 1, 2, 4, 5, 7, 8 ]
integer :: n
do n = 1, 9
block
if (mod(n, 3) == 0) cycle
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 6) then
    print *, "FAIL: more than 6 line(s)"
    stop 1
end if
if ((n) /= vybe_check_w(vybe_check_i)) then
    print *, "FAIL at ", vybe_check_i, " got [", n, "]"
    stop 1
end if
end block
end do
if (vybe_check_i /= 6) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 6"
    stop 1
end if
end program t
