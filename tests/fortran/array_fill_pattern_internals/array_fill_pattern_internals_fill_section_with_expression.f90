! vybe-test: fortran/array_fill_pattern_internals/array_fill_pattern_internals_fill_section_with_expression
! origin: languages/fortran/tests/fortran/test_array_fill_pattern_internals.rs

program array_fill_pattern_internals_fill_section_with_expression
integer :: vybe_check_i = 0
integer :: vybe_check_w(3) = [ 78, 30, 35 ]
    integer, allocatable :: values(:)
    integer :: i
    values = (/ 5, 6, 7, 8 /)
    do i = 2, 3
        values(i) = values(i) * values(1)
    end do
        vybe_check_i = vybe_check_i + 1
    if (vybe_check_i > 3) then
        print *, "FAIL: more than 3 line(s)"
        stop 1
    end if
    if ((sum(values)) /= vybe_check_w(vybe_check_i)) then
        print *, "FAIL at ", vybe_check_i, " got [", sum(values), "]"
        stop 1
    end if
        vybe_check_i = vybe_check_i + 1
    if (vybe_check_i > 3) then
        print *, "FAIL: more than 3 line(s)"
        stop 1
    end if
    if ((values(2)) /= vybe_check_w(vybe_check_i)) then
        print *, "FAIL at ", vybe_check_i, " got [", values(2), "]"
        stop 1
    end if
        vybe_check_i = vybe_check_i + 1
    if (vybe_check_i > 3) then
        print *, "FAIL: more than 3 line(s)"
        stop 1
    end if
    if ((values(3)) /= vybe_check_w(vybe_check_i)) then
        print *, "FAIL at ", vybe_check_i, " got [", values(3), "]"
        stop 1
    end if
if (vybe_check_i /= 3) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 3"
    stop 1
end if
end program array_fill_pattern_internals_fill_section_with_expression
