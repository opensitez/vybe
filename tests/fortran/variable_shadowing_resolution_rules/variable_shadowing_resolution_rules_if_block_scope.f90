! vybe-test: fortran/variable_shadowing_resolution_rules/variable_shadowing_resolution_rules_if_block_scope
! origin: languages/fortran/tests/fortran/test_variable_shadowing_resolution_rules.rs

program variable_shadowing_resolution_rules_if_block_scope
    integer :: level
    level = 3
    if (level > 0) then
        integer :: level
        level = 11
        if ((level) /= 11) then
    print *, "FAIL: want [11] got [", level, "]"
    stop 1
end if
    end if
    if ((level) /= 3) then
    print *, "FAIL: want [3] got [", level, "]"
    stop 1
end if
end program variable_shadowing_resolution_rules_if_block_scope
