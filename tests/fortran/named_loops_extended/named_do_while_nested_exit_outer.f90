! vybe-test: fortran/named_loops_extended/named_do_while_nested_exit_outer
! origin: languages/fortran/tests/fortran/test_named_loops_extended.rs
program t
integer :: vybe_check_i = 0
integer :: vybe_check_w(2) = [ 2, 12 ]
integer :: outer_i, inner_i, total
outer_i = 0
total = 0
outer_loop: do while (outer_i < 5)
outer_i = outer_i + 1
inner: do inner_i = 1, 10
if (inner_i == 3 .and. outer_i == 2) exit outer_loop
total = total + 1
end do inner
end do outer_loop
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 2) then
    print *, "FAIL: more than 2 line(s)"
    stop 1
end if
if ((outer_i) /= vybe_check_w(vybe_check_i)) then
    print *, "FAIL at ", vybe_check_i, " got [", outer_i, "]"
    stop 1
end if
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 2) then
    print *, "FAIL: more than 2 line(s)"
    stop 1
end if
if ((total) /= vybe_check_w(vybe_check_i)) then
    print *, "FAIL at ", vybe_check_i, " got [", total, "]"
    stop 1
end if
if (vybe_check_i /= 2) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 2"
    stop 1
end if
end program t
