! vybe-test: fortran/full_programs/full_program_module_state_and_calling
! origin: languages/fortran/tests/fortran/test_full_programs.rs
module full_program_counter
    integer :: steps = 0
contains
    subroutine advance(by)
        integer, intent(in) :: by
        steps = steps + by
    end subroutine advance

    subroutine reset_counter()
        steps = 0
    end subroutine reset_counter
end module full_program_counter

program full_program_module_state_and_calling
    use full_program_counter
    call advance(4)
    call advance(3)
    if ((steps) /= 7) then
    print *, "FAIL: want [7] got [", steps, "]"
    stop 1
end if
    call reset_counter()
    if ((steps) /= 0) then
    print *, "FAIL: want [0] got [", steps, "]"
    stop 1
end if
end program full_program_module_state_and_calling
