! vybe-test: fortran/do_construct_stop_conditions/test_do_construct_stop_conditions_nested_stop_and_cycle_mix
! origin: languages/fortran/tests/fortran/test_do_construct_stop_conditions.rs

program test_do_construct_stop_conditions
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 12 ]
    integer :: i
    integer :: j
    integer :: total
    total = 0
    do i = 1, 4
        do j = 1, 4
            if (j == 3) cycle
            total = total + 1
            if (j == 4) exit
        end do
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
