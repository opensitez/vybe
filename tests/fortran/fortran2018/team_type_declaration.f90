! vybe-test: fortran/fortran2018/team_type_declaration
! origin: languages/fortran/tests/fortran/test_fortran2018.rs

program test
    use iso_fortran_env
    type(team_type) :: my_team
    print *, 'ok'
end program test
