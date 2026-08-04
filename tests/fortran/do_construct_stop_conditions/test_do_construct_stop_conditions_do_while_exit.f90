! vybe-test: fortran/do_construct_stop_conditions/test_do_construct_stop_conditions_do_while_exit
! origin: languages/fortran/tests/fortran/test_do_construct_stop_conditions.rs

program test_do_construct_stop_conditions
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 15 ]
    integer :: i
    integer :: total
    i = 0
    total = 0
    do while (i < 8)
        i = i + 1
        if (i == 6) exit
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
end program test_do_construct_stop_conditions
