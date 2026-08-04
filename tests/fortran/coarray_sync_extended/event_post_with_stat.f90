! vybe-test: fortran/coarray_sync_extended/event_post_with_stat
! origin: languages/fortran/tests/fortran/test_coarray_sync_extended.rs

program t
    use iso_fortran_env
    type(event_type) :: ev[*]
    integer :: stat
    event post(ev, stat=stat)
    print *, stat
end program t
