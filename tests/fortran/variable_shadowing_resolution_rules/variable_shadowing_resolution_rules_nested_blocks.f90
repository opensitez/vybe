! vybe-test: fortran/variable_shadowing_resolution_rules/variable_shadowing_resolution_rules_nested_blocks
! origin: languages/fortran/tests/fortran/test_variable_shadowing_resolution_rules.rs

program variable_shadowing_resolution_rules_nested_blocks
    integer :: value
    value = 10
    block
        integer :: value
        value = 20
        block
            integer :: value
            value = 30
            if ((value) /= 30) then
    print *, "FAIL: want [30] got [", value, "]"
    stop 1
end if
        end block
        if ((value) /= 20) then
    print *, "FAIL: want [20] got [", value, "]"
    stop 1
end if
    end block
    if ((value) /= 10) then
    print *, "FAIL: want [10] got [", value, "]"
    stop 1
end if
end program variable_shadowing_resolution_rules_nested_blocks
