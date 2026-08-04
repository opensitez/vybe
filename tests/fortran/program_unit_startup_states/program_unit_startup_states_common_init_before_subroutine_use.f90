! vybe-test: fortran/program_unit_startup_states/program_unit_startup_states_common_init_before_subroutine_use
! origin: languages/fortran/tests/fortran/test_program_unit_startup_states.rs

module startup_mod
    integer :: counter = 0
contains
    subroutine bump()
        counter = counter + 1
    end subroutine bump
end module startup_mod

program program_unit_startup_states_common_init_before_subroutine_use
    use startup_mod
    call bump()
    call bump()
    if ((counter) /= 2) then
    print *, "FAIL: want [2] got [", counter, "]"
    stop 1
end if
end program program_unit_startup_states_common_init_before_subroutine_use
