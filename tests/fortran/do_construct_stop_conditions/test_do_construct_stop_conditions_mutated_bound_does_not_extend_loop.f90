! vybe-test: fortran/do_construct_stop_conditions/test_do_construct_stop_conditions_mutated_bound_does_not_extend_loop
! origin: languages/fortran/tests/fortran/test_do_construct_stop_conditions.rs

program test_do_construct_stop_conditions_mutated_bound
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 10 ]
    integer :: i
    integer :: total
    integer :: stop
    stop = 4
    total = 0
    do i = 1, stop
        if (i == 2) stop = 10
        total = total + i
    end do
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
end program test_do_construct_stop_conditions_mutated_bound
