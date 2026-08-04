! vybe-test: fortran/do_construct_reentrancy/test_do_construct_reentrancy_empty_trip_no_iteration
! origin: languages/fortran/tests/fortran/test_do_construct_reentrancy.rs

program test_do_construct_reentrancy_empty_trip
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 0 ]
    integer :: i, sum
    sum = 0
    do i = 1, 5, -1
        sum = sum + i
    end do
        vybe_check_i = vybe_check_i + 1
    if (vybe_check_i > 1) then
        print *, "FAIL: more than 1 line(s)"
        stop 1
    end if
    if ((sum) /= vybe_check_w(vybe_check_i)) then
        print *, "FAIL at ", vybe_check_i, " got [", sum, "]"
        stop 1
    end if
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program test_do_construct_reentrancy_empty_trip
