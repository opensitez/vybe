! vybe-test: fortran/fortran_coverage_bulk/fortran_bulk_associate_array_aliasing
! origin: languages/fortran/tests/fortran/test_fortran_coverage_bulk.rs

program fortran_bulk_associate_array_aliasing
integer :: vybe_check_i = 0
integer :: vybe_check_w(4) = [ 11, 44, 66, 44 ]
    integer :: values(4)
    integer :: i
    values = (/ 11, 22, 33, 44 /)

    associate (middle => values(2:3))
        middle = middle * 2
    end associate

    do i = 1, 4
                vybe_check_i = vybe_check_i + 1
        if (vybe_check_i > 4) then
            print *, "FAIL: more than 4 line(s)"
            stop 1
        end if
        if ((values(i)) /= vybe_check_w(vybe_check_i)) then
            print *, "FAIL at ", vybe_check_i, " got [", values(i), "]"
            stop 1
        end if
    end do
if (vybe_check_i /= 4) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 4"
    stop 1
end if
end program fortran_bulk_associate_array_aliasing
