! vybe-test: fortran/variable_shadowing_resolution_rules/variable_shadowing_resolution_rules_named_block_scope
! origin: languages/fortran/tests/fortran/test_variable_shadowing_resolution_rules.rs

program variable_shadowing_resolution_rules_named_block_scope
    integer :: tally
    tally = 1
    block named_block
        integer :: tally
        tally = 10
        if ((tally) /= 10) then
    print *, "FAIL: want [10] got [", tally, "]"
    stop 1
end if
    end block named_block
    if ((tally) /= 1) then
    print *, "FAIL: want [1] got [", tally, "]"
    stop 1
end if
end program variable_shadowing_resolution_rules_named_block_scope
