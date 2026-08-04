! vybe-test: fortran/coarrays/team_number_in_team
! origin: languages/fortran/tests/fortran/test_coarrays.rs

program test
    use iso_fortran_env
    type(team_type) :: t
    call form_team(1, t)
    change team (t)
        print *, team_number()
    end team
end program test
