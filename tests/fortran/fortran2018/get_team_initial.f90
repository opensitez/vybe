! vybe-test: fortran/fortran2018/get_team_initial
! origin: languages/fortran/tests/fortran/test_fortran2018.rs

program test
    use iso_fortran_env
    type(team_type) :: t
    t = get_team(initial_team)
    print *, 'ok'
end program test
