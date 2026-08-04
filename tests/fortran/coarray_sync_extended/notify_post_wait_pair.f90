! vybe-test: fortran/coarray_sync_extended/notify_post_wait_pair
! origin: languages/fortran/tests/fortran/test_coarray_sync_extended.rs

program t
    use iso_fortran_env
    type(event_type) :: note[*]
    notify post(note[1])
    notify wait(note)
    print *, 'notify done'
end program t
