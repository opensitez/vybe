! vybe-test: fortran/coarray_sync_extended/event_wait_stat_errmsg
! origin: languages/fortran/tests/fortran/test_coarray_sync_extended.rs

program t
    use iso_fortran_env
    type(event_type) :: ev[*]
    integer :: stat
    character(len=80) :: errmsg
    event post(ev)
    event wait(ev, stat=stat, errmsg=errmsg)
    print *, 'waited'
end program t
