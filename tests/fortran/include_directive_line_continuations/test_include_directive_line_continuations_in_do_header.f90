! vybe-test: fortran/include_directive_line_continuations/test_include_directive_line_continuations_in_do_header
! origin: languages/fortran/tests/fortran/test_include_directive_line_continuations.rs

program test_include_directive_line_continuations
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 10 ]
    integer :: s
    s = 0
    do i = 1, 4, &
       1
        s = s + i
    end do
        vybe_check_i = vybe_check_i + 1
    if (vybe_check_i > 1) then
        print *, "FAIL: more than 1 line(s)"
        stop 1
    end if
    if ((s) /= vybe_check_w(vybe_check_i)) then
        print *, "FAIL at ", vybe_check_i, " got [", s, "]"
        stop 1
    end if
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program test_include_directive_line_continuations
