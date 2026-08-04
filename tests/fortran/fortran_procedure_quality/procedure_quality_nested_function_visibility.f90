! vybe-test: fortran/fortran_procedure_quality/procedure_quality_nested_function_visibility
! origin: languages/fortran/tests/fortran/test_fortran_procedure_quality.rs

program procedure_quality_nested_function_visibility
    integer :: result
    call show_add(12, 8, result)
    if ((result) /= 22) then
    print *, "FAIL: want [22] got [", result, "]"
    stop 1
end if

contains
    subroutine show_add(a, b, result)
        integer, intent(in) :: a, b
        integer, intent(out) :: result
        result = inc(a) + inc(b)
    contains
        integer function inc(v)
            integer, intent(in) :: v
            inc = v + 1
        end function inc
    end subroutine show_add
end program procedure_quality_nested_function_visibility
