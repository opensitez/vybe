! vybe-test: fortran/program_unit_startup_states/program_unit_startup_states_module_constants_seed_program_order
! origin: languages/fortran/tests/fortran/test_program_unit_startup_states.rs

module startup_constants
    integer, parameter :: base = 3
    integer :: value
    integer :: computed
    value = base + 7
    computed = value * 2
contains
    integer function get()
        get = value + computed
    end function get
end module startup_constants

program program_unit_startup_states_module_constants_seed_program_order
    use startup_constants
    if ((base) /= 3) then
    print *, "FAIL: want [3] got [", base, "]"
    stop 1
end if
    if ((value) /= 10) then
    print *, "FAIL: want [10] got [", value, "]"
    stop 1
end if
    if ((computed) /= 20) then
    print *, "FAIL: want [20] got [", computed, "]"
    stop 1
end if
    if ((get()) /= 30) then
    print *, "FAIL: want [30] got [", get(), "]"
    stop 1
end if
end program program_unit_startup_states_module_constants_seed_program_order
