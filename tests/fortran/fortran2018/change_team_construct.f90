! vybe-test: fortran/fortran2018/change_team_construct
! origin: languages/fortran/tests/fortran/test_fortran2018.rs

program test
    use iso_fortran_env
    type(team_type) :: t
    call form_team(1, t)
    change team (t)
        print *, this_image()
    end team
end program test
