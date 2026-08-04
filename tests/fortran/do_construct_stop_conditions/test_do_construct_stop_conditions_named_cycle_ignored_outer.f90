! vybe-test: fortran/do_construct_stop_conditions/test_do_construct_stop_conditions_named_cycle_ignored_outer
! origin: languages/fortran/tests/fortran/test_do_construct_stop_conditions.rs

program test_do_construct_stop_conditions
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 8 ]
    integer :: outer
    integer :: inner
    integer :: total
    total = 0
    outer_loop: do outer = 1, 4
        inner_loop: do inner = 1, 4
            if (inner == 3) cycle outer_loop
            total = total + 1
        end do inner_loop
    end do outer_loop
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
end program test_do_construct_stop_conditions
