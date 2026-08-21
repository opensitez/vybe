! vybe-test: fortran/coarray_teams/change_team_construct
! origin: languages/fortran/tests/fortran/test_fortran2018.rs

! `call form_team(1, t)` is NOT Fortran. FORM TEAM is a STATEMENT in F2018,
! spelled `form team (team-number, team-variable)` — no CALL. gfortran left
! `_form_team_` undefined at link because it took the line as a call to an
! external subroutine. With the real statement gfortran compiles AND LINKS it.
program test
    use iso_fortran_env
    type(team_type) :: t
    form team (1, t)
    change team (t)
        print *, this_image()
    end team
end program test
