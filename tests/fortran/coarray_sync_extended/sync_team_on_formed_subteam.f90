! vybe-test: fortran/coarray_sync_extended/sync_team_on_formed_subteam
! origin: languages/fortran/tests/fortran/test_coarray_sync_extended.rs

program t
    use iso_fortran_env
    type(team_type) :: sub
    call form_team(mod(this_image(), 2) + 1, sub)
    sync team (sub)
    print *, 'subteam synced'
end program t
