! vybe-test: fortran/coarray_sync_extended/unlock_with_stat
! origin: languages/fortran/tests/fortran/test_coarray_sync_extended.rs

program t
    use iso_fortran_env
    integer(lock_type) :: lk[*]
    integer :: stat
    lock(lk)
    unlock(lk, stat=stat)
    print *, stat
end program t
