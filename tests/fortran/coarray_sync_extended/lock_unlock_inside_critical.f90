! vybe-test: fortran/coarray_sync_extended/lock_unlock_inside_critical
! origin: languages/fortran/tests/fortran/test_coarray_sync_extended.rs

program t
    use iso_fortran_env
    integer :: tally[*]
    integer(lock_type) :: lk[*]
    tally = 0
    sync all
    critical
        lock(lk[1])
        tally[1] = tally[1] + 1
        unlock(lk[1])
    end critical
    sync all
    if (this_image() == 1) print *, tally
end program t
