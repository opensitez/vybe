! vybe-test: fortran/variable_shadowing_resolution_rules/variable_shadowing_resolution_rules_type_component_hiding
! origin: languages/fortran/tests/fortran/test_variable_shadowing_resolution_rules.rs

program variable_shadowing_resolution_rules_type_component_hiding
    type item
        integer :: value = 1
    end type item
    type(item) :: it
    integer :: value
    value = 9
    it%value = 3
    if ((value) /= 9) then
    print *, "FAIL: want [9] got [", value, "]"
    stop 1
end if
    if ((it%value) /= 3) then
    print *, "FAIL: want [3] got [", it%value, "]"
    stop 1
end if
end program variable_shadowing_resolution_rules_type_component_hiding
