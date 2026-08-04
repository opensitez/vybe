! vybe-test: fortran/subroutine_argument_default_values/subroutine_argument_default_values_array_defaults_with_optional_shape
! origin: languages/fortran/tests/fortran/test_subroutine_argument_default_values.rs

program subroutine_argument_default_values_array_defaults_with_optional_shape
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 6 ]
    integer :: a(3)
    call fill(a)
        vybe_check_i = vybe_check_i + 1
    if (vybe_check_i > 1) then
        print *, "FAIL: more than 1 line(s)"
        stop 1
    end if
    if ((sum(a)) /= vybe_check_w(vybe_check_i)) then
        print *, "FAIL at ", vybe_check_i, " got [", sum(a), "]"
        stop 1
    end if
contains
    subroutine fill(out, scale)
        integer, intent(out) :: out(:)
        integer, intent(in), optional :: scale
        integer :: i
        integer :: factor
        if (present(scale)) then
            factor = scale
        else
            factor = 1
        end if
        do i = 1, size(out)
            out(i) = i * factor
        end do
    end subroutine fill
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program subroutine_argument_default_values_array_defaults_with_optional_shape
