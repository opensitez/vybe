! vybe-test: fortran/program_unit_startup_states/program_unit_startup_states_block_data_like_seed
! origin: languages/fortran/tests/fortran/test_program_unit_startup_states.rs

module block_state
    integer :: a
    integer :: b
    integer, save :: token = 11
    contains
        subroutine init()
            a = 5
            b = a + token
        end subroutine init
end module block_state

program program_unit_startup_states_block_data_like_seed
    use block_state
    call init()
    if ((a) /= 5) then
    print *, "FAIL: want [5] got [", a, "]"
    stop 1
end if
    if ((b) /= 16) then
    print *, "FAIL: want [16] got [", b, "]"
    stop 1
end if
    if ((token) /= 11) then
    print *, "FAIL: want [11] got [", token, "]"
    stop 1
end if
end program program_unit_startup_states_block_data_like_seed
