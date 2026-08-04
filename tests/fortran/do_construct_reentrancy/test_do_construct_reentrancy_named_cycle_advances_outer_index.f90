! vybe-test: fortran/do_construct_reentrancy/test_do_construct_reentrancy_named_cycle_advances_outer_index
! origin: languages/fortran/tests/fortran/test_do_construct_reentrancy.rs

program test_do_construct_reentrancy_named_cycle
integer :: vybe_check_i = 0
integer :: vybe_check_w(3) = [ 6, 5, 2 ]
    integer :: outer
    integer :: inner
    integer :: total
    total = 0
    outer_loop: do outer = 1, 4
        do inner = 1, 4
            if (mod(inner, 2) == 0) cycle outer_loop
            total = total + outer
        end do
    end do outer_loop
        vybe_check_i = vybe_check_i + 1
    if (vybe_check_i > 3) then
        print *, "FAIL: more than 3 line(s)"
        stop 1
    end if
    if ((total) /= vybe_check_w(vybe_check_i)) then
        print *, "FAIL at ", vybe_check_i, " got [", total, "]"
        stop 1
    end if
        vybe_check_i = vybe_check_i + 1
    if (vybe_check_i > 3) then
        print *, "FAIL: more than 3 line(s)"
        stop 1
    end if
    if ((outer) /= vybe_check_w(vybe_check_i)) then
        print *, "FAIL at ", vybe_check_i, " got [", outer, "]"
        stop 1
    end if
        vybe_check_i = vybe_check_i + 1
    if (vybe_check_i > 3) then
        print *, "FAIL: more than 3 line(s)"
        stop 1
    end if
    if ((inner) /= vybe_check_w(vybe_check_i)) then
        print *, "FAIL at ", vybe_check_i, " got [", inner, "]"
        stop 1
    end if
if (vybe_check_i /= 3) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 3"
    stop 1
end if
end program test_do_construct_reentrancy_named_cycle
