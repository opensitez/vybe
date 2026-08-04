! vybe-test: fortran/coarrays/lock_unlock_coarray
! origin: languages/fortran/tests/fortran/test_coarrays.rs

program test
    use iso_fortran_env
    integer(lock_type) :: lk[*]
    lock(lk[1])
    print *, 'locked'
    unlock(lk[1])
    print *, 'unlocked'
end program test
