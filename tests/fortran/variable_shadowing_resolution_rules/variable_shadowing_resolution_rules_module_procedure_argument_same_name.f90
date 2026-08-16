! vybe-test: fortran/variable_shadowing_resolution_rules/variable_shadowing_resolution_rules_module_procedure_argument_same_name
! origin: languages/fortran/tests/fortran/test_variable_shadowing_resolution_rules.rs

module shadow_args_mod
    integer :: value = 13
contains
    subroutine emit(value)
        integer, intent(in) :: value
        if ((value) /= 99) then
    print *, "FAIL: want [99] got [", value, "]"
    stop 1
end if
    end subroutine emit
end module shadow_args_mod

program t
    use shadow_args_mod
    call emit(99)
    if ((value) /= 13) then
    print *, "FAIL: want [13] got [", value, "]"
    stop 1
end if
end program t
