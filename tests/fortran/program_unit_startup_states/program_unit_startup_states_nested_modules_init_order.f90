! vybe-test: fortran/program_unit_startup_states/program_unit_startup_states_nested_modules_init_order
! origin: languages/fortran/tests/fortran/test_program_unit_startup_states.rs

module alpha
    integer, parameter :: p = 2
end module alpha

module beta
    use alpha
    integer, parameter :: q = p + 3
end module beta

program program_unit_startup_states_nested_modules_init_order
    use beta
    if ((p) /= 2) then
    print *, "FAIL: want [2] got [", p, "]"
    stop 1
end if
    if ((q) /= 5) then
    print *, "FAIL: want [5] got [", q, "]"
    stop 1
end if
end program program_unit_startup_states_nested_modules_init_order
