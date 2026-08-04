! vybe-test: fortran/program_unit_startup_states/program_unit_startup_states_derived_type_default_init
! origin: languages/fortran/tests/fortran/test_program_unit_startup_states.rs

module startup_types
    type cfg
        integer :: a = 1
        integer :: b = 2
        integer :: c = a + b
    end type cfg
end module startup_types

program program_unit_startup_states_derived_type_default_init
    use startup_types
    type(cfg) :: item
    if ((item%a) /= 1) then
    print *, "FAIL: want [1] got [", item%a, "]"
    stop 1
end if
    if ((item%b) /= 2) then
    print *, "FAIL: want [2] got [", item%b, "]"
    stop 1
end if
    if ((item%c) /= 3) then
    print *, "FAIL: want [3] got [", item%c, "]"
    stop 1
end if
end program program_unit_startup_states_derived_type_default_init
