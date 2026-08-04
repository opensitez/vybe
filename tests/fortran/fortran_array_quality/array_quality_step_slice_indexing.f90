! vybe-test: fortran/fortran_array_quality/array_quality_step_slice_indexing
! origin: languages/fortran/tests/fortran/test_fortran_array_quality.rs

program array_quality_step_slice_indexing
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 16 ]
    integer, dimension(8) :: values
    integer :: i, total
    values = (/ 1, 2, 3, 4, 5, 6, 7, 8 /)
    total = 0
    do i = 1, 8, 2
        total = total + values(i)
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
end program array_quality_step_slice_indexing
