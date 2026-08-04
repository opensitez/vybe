! vybe-test: fortran/program_unit_startup_states/program_unit_startup_states_initialization_dependencies
! origin: languages/fortran/tests/fortran/test_program_unit_startup_states.rs

program program_unit_startup_states_initialization_dependencies
    integer :: first = 4
    integer :: second = first + 1
    integer :: third = second + 1
    if ((first) /= 4) then
    print *, "FAIL: want [4] got [", first, "]"
    stop 1
end if
    if ((second) /= 5) then
    print *, "FAIL: want [5] got [", second, "]"
    stop 1
end if
    if ((third) /= 6) then
    print *, "FAIL: want [6] got [", third, "]"
    stop 1
end if
end program program_unit_startup_states_initialization_dependencies
