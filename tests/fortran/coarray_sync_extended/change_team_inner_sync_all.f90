! vybe-test: fortran/coarray_sync_extended/change_team_inner_sync_all
! origin: languages/fortran/tests/fortran/test_coarray_sync_extended.rs

program t
    use iso_fortran_env
    type(team_type) :: t
    call form_team(1, t)
    change team (t)
        sync all
        print *, team_number()
    end team
end program t
