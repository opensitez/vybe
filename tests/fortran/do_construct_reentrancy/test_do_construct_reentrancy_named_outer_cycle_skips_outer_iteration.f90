! vybe-test: fortran/do_construct_reentrancy/test_do_construct_reentrancy_named_outer_cycle_skips_outer_iteration
! origin: languages/fortran/tests/fortran/test_do_construct_reentrancy.rs

program test_do_construct_reentrancy_named_outer_cycle
integer :: vybe_check_i = 0
integer :: vybe_check_w(2) = [ 5, 10 ]
    integer :: outer, inner, total
    total = 0
    outer_loop: do outer = 1, 4
        do inner = 1, 3
            if (inner == 2 .and. outer == 2) cycle outer_loop
            total = total + 1
        end do
    end do outer_loop
        vybe_check_i = vybe_check_i + 1
    if (vybe_check_i > 2) then
        print *, "FAIL: more than 2 line(s)"
        stop 1
    end if
    if ((outer) /= vybe_check_w(vybe_check_i)) then
        print *, "FAIL at ", vybe_check_i, " got [", outer, "]"
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
end program test_do_construct_reentrancy_named_outer_cycle
