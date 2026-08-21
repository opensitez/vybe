! vybe-test: fortran/coarray_teams/team_number_intrinsic
! origin: languages/fortran/tests/fortran/test_fortran2018.rs

program test
    use iso_fortran_env
    type(team_type) :: t
    t = get_team()
    print *, team_number(t)
end program test
