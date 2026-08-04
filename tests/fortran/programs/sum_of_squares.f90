! vybe-test: fortran/programs/sum_of_squares
! origin: languages/fortran/tests/fortran/test_programs.rs

program test
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 385 ]
    integer :: i, total
    total = 0
    do i = 1, 10
        total = total + i * i
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
end program test
