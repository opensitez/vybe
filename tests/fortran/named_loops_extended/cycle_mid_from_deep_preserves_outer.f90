! vybe-test: fortran/named_loops_extended/cycle_mid_from_deep_preserves_outer
! origin: languages/fortran/tests/fortran/test_named_loops_extended.rs
program t
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 24 ]
integer :: i, j, k, total
total = 0
outer: do i = 1, 3
mid: do j = 1, 3
inner: do k = 1, 3
if (j == 2 .and. k == 1) cycle mid
total = total + 1
end do inner
end do mid
end do outer
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 1) then
    print *, "FAIL: more than 1 line(s)"
    stop 1
end if
if ((total) /= vybe_check_w(vybe_check_i)) then
    print *, "FAIL at ", vybe_check_i, " got [", total, "]"
    stop 1
end if
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program t
