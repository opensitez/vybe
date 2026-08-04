! vybe-test: fortran/fortran_procedure_quality/procedure_quality_pure_function
! origin: languages/fortran/tests/fortran/test_fortran_procedure_quality.rs

program procedure_quality_pure_function
    integer :: output
    output = negate_if_negative(-3)
    if ((output) /= 3) then
    print *, "FAIL: want [3] got [", output, "]"
    stop 1
end if

contains
    integer function negate_if_negative(v)
        integer, intent(in) :: v
        if (v < 0) negate_if_negative = -v
        if (v >= 0) negate_if_negative = v
    end function negate_if_negative
end program procedure_quality_pure_function
