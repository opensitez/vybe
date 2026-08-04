! vybe-test: fortran/coarray_sync_extended/lock_errmsg_when_busy
! origin: languages/fortran/tests/fortran/test_coarray_sync_extended.rs

program t
    use iso_fortran_env
    integer(lock_type) :: lk[*]
    integer :: stat
    character(len=80) :: errmsg
    lock(lk, stat=stat, errmsg=errmsg, acquired_lock=.true.)
    if (stat == 0) unlock(lk)
    print *, trim(errmsg)
end program t
