! vybe-test: fortran/program_unit_startup_states/program_unit_startup_states_data_initialization_before_block
! origin: languages/fortran/tests/fortran/test_program_unit_startup_states.rs

program program_unit_startup_states_data_initialization_before_block
    integer :: seed
    data seed /99/
    block
        integer :: doubled
        doubled = seed * 2
        if ((seed) /= 99) then
    print *, "FAIL: want [99] got [", seed, "]"
    stop 1
end if
        if ((doubled) /= 198) then
    print *, "FAIL: want [198] got [", doubled, "]"
    stop 1
end if
    end block
end program program_unit_startup_states_data_initialization_before_block
