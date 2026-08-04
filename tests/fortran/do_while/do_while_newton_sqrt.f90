! vybe-test: fortran/do_while/do_while_newton_sqrt
! origin: languages/fortran/tests/fortran/test_do_while.rs

program test
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 1 ]
    real :: x = 2.0, g = 1.0, prev
    integer :: it
    it = 0
    do while (.true.)
        prev = g
        g = 0.5 * (g + x / g)
        it = it + 1
        if (abs(g - prev) < 1e-7 .or. it > 20) exit
    end do
    if (abs(g * g - x) < 1.0e-3) then
                vybe_check_i = vybe_check_i + 1
        if (vybe_check_i > 1) then
            print *, "FAIL: more than 1 line(s)"
            stop 1
        end if
        if ((1) /= vybe_check_w(vybe_check_i)) then
            print *, "FAIL at ", vybe_check_i, " got [", 1, "]"
            stop 1
        end if
    else
                vybe_check_i = vybe_check_i + 1
        if (vybe_check_i > 1) then
            print *, "FAIL: more than 1 line(s)"
            stop 1
        end if
        if ((0) /= vybe_check_w(vybe_check_i)) then
            print *, "FAIL at ", vybe_check_i, " got [", 0, "]"
            stop 1
        end if
    end if
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program test
