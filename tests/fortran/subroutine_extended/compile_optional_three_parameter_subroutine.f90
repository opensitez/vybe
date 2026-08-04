! vybe-test: fortran/subroutine_extended/compile_optional_three_parameter_subroutine
! origin: languages/fortran/tests/fortran/test_subroutine_extended.rs

program t
integer :: vybe_check_i = 0
integer :: vybe_check_w(3) = [ 1, 3, 6 ]
    call tagged(1)
    call tagged(1, 2)
    call tagged(1, 2, 3)
contains
    subroutine tagged(a, b, c)
        integer, intent(in) :: a
        integer, intent(in), optional :: b, c
        integer :: total
        total = a
        if (present(b)) total = total + b
        if (present(c)) total = total + c
                vybe_check_i = vybe_check_i + 1
        if (vybe_check_i > 3) then
            print *, "FAIL: more than 3 line(s)"
            stop 1
        end if
        if ((total) /= vybe_check_w(vybe_check_i)) then
            print *, "FAIL at ", vybe_check_i, " got [", total, "]"
            stop 1
        end if
    end subroutine tagged
if (vybe_check_i /= 3) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 3"
    stop 1
end if
end program t
