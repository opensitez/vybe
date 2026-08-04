! vybe-test: fortran/fortran_array_quality/array_quality_transformation_count
! origin: languages/fortran/tests/fortran/test_fortran_array_quality.rs

program array_quality_transformation_count
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 3 ]
    integer, dimension(6) :: values
    integer :: i
    integer :: zeros
    values = (/ 0, 1, 0, 1, 0, 1 /)
    zeros = 0
    do i = 1, 6
        if (values(i) == 0) zeros = zeros + 1
    end do
        vybe_check_i = vybe_check_i + 1
    if (vybe_check_i > 1) then
        print *, "FAIL: more than 1 line(s)"
        stop 1
    end if
    if ((zeros) /= vybe_check_w(vybe_check_i)) then
        print *, "FAIL at ", vybe_check_i, " got [", zeros, "]"
        stop 1
    end if
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program array_quality_transformation_count
