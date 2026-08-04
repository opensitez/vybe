! vybe-test: fortran/fortran_procedure_quality/procedure_quality_internal_state_update
! origin: languages/fortran/tests/fortran/test_fortran_procedure_quality.rs

module counter_module
    integer, save :: counter = 0

    contains
    subroutine bump()
        counter = counter + 1
    end subroutine bump

    integer function value()
        value = counter
    end function value
end module counter_module

program procedure_quality_internal_state_update
    use counter_module
    call bump()
    call bump()
    if ((value()) /= 2) then
    print *, "FAIL: want [2] got [", value(), "]"
    stop 1
end if
end program procedure_quality_internal_state_update
