! vybe-test: fortran/fortran_procedure_quality/procedure_quality_function_return
! origin: languages/fortran/tests/fortran/test_fortran_procedure_quality.rs

program procedure_quality_function_return
    integer :: value

    value = square(7)
    if ((value) /= 49) then
    print *, "FAIL: want [49] got [", value, "]"
    stop 1
end if

contains

integer function square(x)
    integer, intent(in) :: x
    square = x * x
end function square
end program procedure_quality_function_return
