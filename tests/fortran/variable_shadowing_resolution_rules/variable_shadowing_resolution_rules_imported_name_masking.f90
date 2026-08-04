! vybe-test: fortran/variable_shadowing_resolution_rules/variable_shadowing_resolution_rules_imported_name_masking
! origin: languages/fortran/tests/fortran/test_variable_shadowing_resolution_rules.rs

module shadow_a
    integer :: token = 1
end module shadow_a

program variable_shadowing_resolution_rules_imported_name_masking
    use shadow_a, only: module_token => token
    integer :: token
    token = 9
    if ((token) /= 9) then
    print *, "FAIL: want [9] got [", token, "]"
    stop 1
end if
    if ((module_token) /= 1) then
    print *, "FAIL: want [1] got [", module_token, "]"
    stop 1
end if
end program variable_shadowing_resolution_rules_imported_name_masking
