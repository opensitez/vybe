! vybe-test: fortran/array_masked_array_operations/array_masked_array_operations_where_mask_with_stride
! origin: languages/fortran/tests/fortran/test_array_masked_array_operations.rs

program array_masked_array_operations_where_mask_with_stride
integer :: vybe_check_i = 0
integer :: vybe_check_w(4) = [ 46, 1, 10, 30 ]
    integer :: values(6)
    integer :: result(6)
    integer :: i
    values = (/ 1, 2, 3, 4, 5, 6 /)
    do i = 1, 6
        if (mod(i, 2) == 0) then
            where (values(i) > 0)
                result(i) = values(i) * 5
            end where
        else
            result(i) = values(i)
        end if
    end do
        vybe_check_i = vybe_check_i + 1
    if (vybe_check_i > 4) then
        print *, "FAIL: more than 4 line(s)"
        stop 1
    end if
    if ((sum(result)) /= vybe_check_w(vybe_check_i)) then
        print *, "FAIL at ", vybe_check_i, " got [", sum(result), "]"
        stop 1
    end if
        vybe_check_i = vybe_check_i + 1
    if (vybe_check_i > 4) then
        print *, "FAIL: more than 4 line(s)"
        stop 1
    end if
    if ((result(1)) /= vybe_check_w(vybe_check_i)) then
        print *, "FAIL at ", vybe_check_i, " got [", result(1), "]"
        stop 1
    end if
        vybe_check_i = vybe_check_i + 1
    if (vybe_check_i > 4) then
        print *, "FAIL: more than 4 line(s)"
        stop 1
    end if
    if ((result(2)) /= vybe_check_w(vybe_check_i)) then
        print *, "FAIL at ", vybe_check_i, " got [", result(2), "]"
        stop 1
    end if
        vybe_check_i = vybe_check_i + 1
    if (vybe_check_i > 4) then
        print *, "FAIL: more than 4 line(s)"
        stop 1
    end if
    if ((result(6)) /= vybe_check_w(vybe_check_i)) then
        print *, "FAIL at ", vybe_check_i, " got [", result(6), "]"
        stop 1
    end if
if (vybe_check_i /= 4) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 4"
    stop 1
end if
end program array_masked_array_operations_where_mask_with_stride
