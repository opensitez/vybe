! vybe-test: fortran/coarrays/form_team_and_change
! origin: languages/fortran/tests/fortran/test_coarrays.rs

program test
    use iso_fortran_env
    type(team_type) :: odd_even
    integer :: color
    color = mod(this_image(), 2) + 1
    call form_team(color, odd_even)
    change team (odd_even)
        print *, this_image(), 'in subteam', team_number()
    end team
    sync all
    if (this_image() == 1) print *, 'done'
end program test
