! vybe-test: fortran/program_unit_startup_states/program_unit_startup_states_block_rebind_seed
! origin: languages/fortran/tests/fortran/test_program_unit_startup_states.rs

program program_unit_startup_states_block_rebind_seed
    integer :: count
    count = 1
    block
        integer :: count
        count = 99
        if ((count) /= 99) then
    print *, "FAIL: want [99] got [", count, "]"
    stop 1
end if
    end block
    if ((count) /= 1) then
    print *, "FAIL: want [1] got [", count, "]"
    stop 1
end if
end program program_unit_startup_states_block_rebind_seed
