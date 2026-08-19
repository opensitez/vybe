! vybe-test: fortran/array_fill_pattern_internals/array_fill_pattern_internals_fill_with_conditional_offsets
! origin: languages/fortran/tests/fortran/test_array_fill_pattern_internals.rs

program array_fill_pattern_internals_fill_with_conditional_offsets
integer :: vybe_check_i = 0
integer :: vybe_check_w(3) = [ 42, 16, 58 ]
    integer, allocatable :: values(:)
    integer :: i
    values = (/ 1, 2, 3, 4, 5, 6, 7 /)
    do i = lbound(values,1), ubound(values,1)
        if (mod(i,2) == 0) values(i) = values(i) + 10
    end do
        vybe_check_i = vybe_check_i + 1
    if (vybe_check_i > 3) then
        print *, "FAIL: more than 3 line(s)"
        stop 1
    end if
    if ((values(2) + values(4) + values(6)) /= vybe_check_w(vybe_check_i)) then
        print *, "FAIL at ", vybe_check_i, " got [", values(2) + values(4) + values(6), "]"
        stop 1
    end if
        vybe_check_i = vybe_check_i + 1
    if (vybe_check_i > 3) then
        print *, "FAIL: more than 3 line(s)"
        stop 1
    end if
    if ((values(1) + values(3) + values(5) + values(7)) /= vybe_check_w(vybe_check_i)) then
        print *, "FAIL at ", vybe_check_i, " got [", values(1) + values(3) + values(5) + values(7), "]"
        stop 1
    end if
        vybe_check_i = vybe_check_i + 1
    if (vybe_check_i > 3) then
        print *, "FAIL: more than 3 line(s)"
        stop 1
    end if
    if ((sum(values)) /= vybe_check_w(vybe_check_i)) then
        print *, "FAIL at ", vybe_check_i, " got [", sum(values), "]"
        stop 1
    end if
if (vybe_check_i /= 3) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 3"
    stop 1
end if
end program array_fill_pattern_internals_fill_with_conditional_offsets
