! vybe-test: fortran/fortran_procedure_quality/procedure_quality_optional_argument_default
! origin: languages/fortran/tests/fortran/test_fortran_procedure_quality.rs

program procedure_quality_optional_argument_default
    integer :: out_a
    integer :: out_b
    call multiply(3, out_a)
    call multiply(3, out_b, 4)
    if ((out_a) /= 6) then
    print *, "FAIL: want [6] got [", out_a, "]"
    stop 1
end if
    if ((out_b) /= 12) then
    print *, "FAIL: want [12] got [", out_b, "]"
    stop 1
end if

contains
    subroutine multiply(a, result, scale)
        integer, intent(in) :: a
        integer, intent(out) :: result
        integer, intent(in), optional :: scale
        if (present(scale)) then
            result = a * scale
        else
            result = a * 2
        end if
    end subroutine multiply
end program procedure_quality_optional_argument_default
