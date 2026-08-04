! vybe-test: fortran/do_while/nested_do_while_with_exit
! origin: languages/fortran/tests/fortran/test_do_while.rs

program test
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 10 ]
    integer :: i = 0, j, count
    count = 0
    do while (i < 5)
        i = i + 1
        j = 0
        do while (j < 5)
            j = j + 1
            if (j == 3) exit
            count = count + 1
        end do
    end do
        vybe_check_i = vybe_check_i + 1
    if (vybe_check_i > 1) then
        print *, "FAIL: more than 1 line(s)"
        stop 1
    end if
    if ((count) /= vybe_check_w(vybe_check_i)) then
        print *, "FAIL at ", vybe_check_i, " got [", count, "]"
        stop 1
    end if
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program test
