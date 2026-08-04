! vybe-test: fortran/fortran_procedure_quality/procedure_quality_interface_specificity
! origin: languages/fortran/tests/fortran/test_fortran_procedure_quality.rs

program procedure_quality_interface_specificity
    integer :: int_out
    call call_increment(int_out)
    if ((int_out) /= 7) then
    print *, "FAIL: want [7] got [", int_out, "]"
    stop 1
end if

contains
    interface
        function int_add_one(v) result(r)
            integer :: r
            integer, intent(in) :: v
        end function int_add_one
    end interface

    function int_add_one(v) result(r)
        integer, intent(in) :: v
        integer :: r
        r = v + 1
    end function int_add_one

    subroutine call_increment(result)
        integer, intent(out) :: result
        result = int_add_one(6)
    end subroutine call_increment
end program procedure_quality_interface_specificity
