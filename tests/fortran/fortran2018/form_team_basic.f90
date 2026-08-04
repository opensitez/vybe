! vybe-test: fortran/fortran2018/form_team_basic
! origin: languages/fortran/tests/fortran/test_fortran2018.rs

program test
    use iso_fortran_env
    type(team_type) :: t
    call form_team(1, t)
    print *, 'ok'
end program test
