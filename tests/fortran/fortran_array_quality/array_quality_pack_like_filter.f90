! vybe-test: fortran/fortran_array_quality/array_quality_pack_like_filter
! origin: languages/fortran/tests/fortran/test_fortran_array_quality.rs

program array_quality_pack_like_filter
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 6 ]
    integer, dimension(5) :: input
    integer :: i
    integer :: total
    input = (/ 1, 0, 2, 0, 3 /)
    total = 0
    do i = 1, 5
        if (input(i) > 0) total = total + input(i)
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
end program array_quality_pack_like_filter
