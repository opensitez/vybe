! vybe-test: fortran/named_loops_extended/exit_outer_from_triple_when_product_is_twentyfour
! origin: languages/fortran/tests/fortran/test_named_loops_extended.rs
program t
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 24 ]
integer :: i, j, k
found: do i = 1, 6
mid: do j = 1, 6
deep: do k = 1, 6
if (i * j * k == 24) exit found
end do deep
end do mid
end do found
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 1) then
    print *, "FAIL: more than 1 line(s)"
    stop 1
end if
if ((i * j * k) /= vybe_check_w(vybe_check_i)) then
    print *, "FAIL at ", vybe_check_i, " got [", i * j * k, "]"
    stop 1
end if
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program t
