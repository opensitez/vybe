! vybe-test: fortran/array_masked_array_operations/array_masked_array_operations_masked_update_inside_do
! origin: languages/fortran/tests/fortran/test_array_masked_array_operations.rs

program array_masked_array_operations_masked_update_inside_do
integer :: vybe_check_i = 0
integer :: vybe_check_w(3) = [ 39, 5, 15 ]
    integer :: values(4)
    integer :: result(4)
    integer :: i
    values = (/ 4, 8, 12, 16 /)
    do i = 1, 4
        if (i <= 2) then
            result(i) = values(i) + 1
        else
            result(i) = values(i) - 1
        end if
    end do
        vybe_check_i = vybe_check_i + 1
    if (vybe_check_i > 3) then
        print *, "FAIL: more than 3 line(s)"
        stop 1
    end if
    if ((sum(result)) /= vybe_check_w(vybe_check_i)) then
        print *, "FAIL at ", vybe_check_i, " got [", sum(result), "]"
        stop 1
    end if
        vybe_check_i = vybe_check_i + 1
    if (vybe_check_i > 3) then
        print *, "FAIL: more than 3 line(s)"
        stop 1
    end if
    if ((result(1)) /= vybe_check_w(vybe_check_i)) then
        print *, "FAIL at ", vybe_check_i, " got [", result(1), "]"
        stop 1
    end if
        vybe_check_i = vybe_check_i + 1
    if (vybe_check_i > 3) then
        print *, "FAIL: more than 3 line(s)"
        stop 1
    end if
    if ((result(4)) /= vybe_check_w(vybe_check_i)) then
        print *, "FAIL at ", vybe_check_i, " got [", result(4), "]"
        stop 1
    end if
if (vybe_check_i /= 3) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 3"
    stop 1
end if
end program array_masked_array_operations_masked_update_inside_do
