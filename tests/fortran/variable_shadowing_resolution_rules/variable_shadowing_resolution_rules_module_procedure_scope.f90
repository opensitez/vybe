! vybe-test: fortran/variable_shadowing_resolution_rules/variable_shadowing_resolution_rules_module_procedure_scope
! origin: languages/fortran/tests/fortran/test_variable_shadowing_resolution_rules.rs

module scope_rules_mod
    integer :: token = 1
contains
    subroutine report()
        integer :: token
        token = 4
        if ((token) /= 1) then
    print *, "FAIL: want [1] got [", token, "]"
    stop 1
end if
    end subroutine report
end module scope_rules_mod

program variable_shadowing_resolution_rules_module_procedure_scope
    use scope_rules_mod
    if ((token) /= 4) then
    print *, "FAIL: want [4] got [", token, "]"
    stop 1
end if
    call report()
    if ((token) /= 1) then
    print *, "FAIL: want [1] got [", token, "]"
    stop 1
end if
end program variable_shadowing_resolution_rules_module_procedure_scope
