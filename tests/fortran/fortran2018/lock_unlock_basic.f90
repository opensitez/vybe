! vybe-test: fortran/fortran2018/lock_unlock_basic
! origin: languages/fortran/tests/fortran/test_fortran2018.rs

program test
    use iso_fortran_env
    integer(lock_type) :: lk[*]
    lock(lk)
    print *, 'locked'
    unlock(lk)
    print *, 'unlocked'
end program test
