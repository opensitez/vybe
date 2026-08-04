! vybe-test: fortran/coarrays/sync_team_basic
! origin: languages/fortran/tests/fortran/test_coarrays.rs

program test
    use iso_fortran_env
    type(team_type) :: t
    t = get_team()
    sync team (t)
    print *, 'ok'
end program test
