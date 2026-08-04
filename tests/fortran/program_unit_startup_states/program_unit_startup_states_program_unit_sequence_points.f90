! vybe-test: fortran/program_unit_startup_states/program_unit_startup_states_program_unit_sequence_points
! origin: languages/fortran/tests/fortran/test_program_unit_startup_states.rs

program program_unit_startup_states_program_unit_sequence_points
    integer :: x = 1
    integer :: y = 2
    integer :: z
    z = x + y
    if ((x) /= 1) then
    print *, "FAIL: want [1] got [", x, "]"
    stop 1
end if
    if ((y) /= 2) then
    print *, "FAIL: want [2] got [", y, "]"
    stop 1
end if
    if ((z) /= 3) then
    print *, "FAIL: want [3] got [", z, "]"
    stop 1
end if
end program program_unit_startup_states_program_unit_sequence_points
