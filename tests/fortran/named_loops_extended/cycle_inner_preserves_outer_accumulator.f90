! vybe-test: fortran/named_loops_extended/cycle_inner_preserves_outer_accumulator
! origin: languages/fortran/tests/fortran/test_named_loops_extended.rs
program t
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 54 ]
integer :: i, j, sum
sum = 0
outer: do i = 1, 3
inner: do j = 1, 4
if (j == 1) cycle inner
sum = sum + i * j
end do inner
end do outer
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 1) then
    print *, "FAIL: more than 1 line(s)"
    stop 1
end if
if ((sum) /= vybe_check_w(vybe_check_i)) then
    print *, "FAIL at ", vybe_check_i, " got [", sum, "]"
    stop 1
end if
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program t
