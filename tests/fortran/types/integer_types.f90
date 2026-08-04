! vybe-test: fortran/types/integer_types
! origin: languages/fortran/tests/fortran/test_types.rs

program test
integer :: vybe_check_i = 0
real :: vybe_check_w(2) = [ 10, 3.14 ]
    integer :: a = 10
    real :: b = 3.14
    double precision :: c = 2.718281828
    logical :: d = .true.
    character(len=10) :: e = "hello"
        vybe_check_i = vybe_check_i + 1
    if (vybe_check_i > 2) then
        print *, "FAIL: more than 2 line(s)"
        stop 1
    end if
    if (abs((a) - vybe_check_w(vybe_check_i)) > 1.0e-6) then
        print *, "FAIL at ", vybe_check_i, " got [", a, "]"
        stop 1
    end if
        vybe_check_i = vybe_check_i + 1
    if (vybe_check_i > 2) then
        print *, "FAIL: more than 2 line(s)"
        stop 1
    end if
    if (abs((b) - vybe_check_w(vybe_check_i)) > 1.0e-6) then
        print *, "FAIL at ", vybe_check_i, " got [", b, "]"
        stop 1
    end if
        vybe_check_i = vybe_check_i + 1
    if (vybe_check_i > 2) then
        print *, "FAIL: more than 2 line(s)"
        stop 1
    end if
    if (abs((d) - vybe_check_w(vybe_check_i)) > 1.0e-6) then
        print *, "FAIL at ", vybe_check_i, " got [", d, "]"
        stop 1
    end if
        vybe_check_i = vybe_check_i + 1
    if (vybe_check_i > 2) then
        print *, "FAIL: more than 2 line(s)"
        stop 1
    end if
    if (abs((e) - vybe_check_w(vybe_check_i)) > 1.0e-6) then
        print *, "FAIL at ", vybe_check_i, " got [", e, "]"
        stop 1
    end if
if (vybe_check_i /= 2) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 2"
    stop 1
end if
end program test
