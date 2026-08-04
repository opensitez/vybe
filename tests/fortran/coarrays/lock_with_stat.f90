! vybe-test: fortran/coarrays/lock_with_stat
! origin: languages/fortran/tests/fortran/test_coarrays.rs

program test
    use iso_fortran_env
    integer(lock_type) :: lk[*]
    integer :: stat
    lock(lk, stat=stat, acquired_lock=.true.)
    if (stat == 0) then
        print *, 'locked ok'
        unlock(lk)
    end if
end program test
