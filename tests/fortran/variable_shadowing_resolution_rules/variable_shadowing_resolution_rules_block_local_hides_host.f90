! vybe-test: fortran/variable_shadowing_resolution_rules/variable_shadowing_resolution_rules_block_local_hides_host
! origin: languages/fortran/tests/fortran/test_variable_shadowing_resolution_rules.rs

program variable_shadowing_resolution_rules_block_local_hides_host
    integer :: x
    x = 1
    block
        integer :: x
        x = 2
        if ((x) /= 2) then
    print *, "FAIL: want [2] got [", x, "]"
    stop 1
end if
    end block
    if ((x) /= 1) then
    print *, "FAIL: want [1] got [", x, "]"
    stop 1
end if
end program variable_shadowing_resolution_rules_block_local_hides_host
