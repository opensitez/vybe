! vybe-test: fortran/do_construct_stop_conditions/test_do_construct_stop_conditions_zero_step_nested_no_progress_guarded_by_if
! origin: languages/fortran/tests/fortran/test_do_construct_stop_conditions.rs

program test_do_construct_stop_conditions
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 0 ]
    integer :: i
    integer :: total
    total = 0
    i = 1
    do while (i < 2)
        if (i == 1) then
            i = i + 1
            cycle
        end if
        total = total + 1
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
end program test_do_construct_stop_conditions
