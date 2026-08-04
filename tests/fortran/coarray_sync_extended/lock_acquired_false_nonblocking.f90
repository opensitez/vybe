! vybe-test: fortran/coarray_sync_extended/lock_acquired_false_nonblocking
! origin: languages/fortran/tests/fortran/test_coarray_sync_extended.rs

program t
    use iso_fortran_env
    integer(lock_type) :: lk[*]
    integer :: stat
    lock(lk, stat=stat, acquired_lock=.false.)
    print *, stat
end program t
