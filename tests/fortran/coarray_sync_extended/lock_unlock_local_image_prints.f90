! vybe-test: fortran/coarray_sync_extended/lock_unlock_local_image_prints
! origin: languages/fortran/tests/fortran/test_coarray_sync_extended.rs
program t
use iso_fortran_env
integer(lock_type) :: lk[*]
lock(lk)
if (trim('held') /= "held") then
    print *, "FAIL: want [held] got [", 'held', "]"
    stop 1
end if
unlock(lk)
if (trim('free') /= "free") then
    print *, "FAIL: want [free] got [", 'free', "]"
    stop 1
end if
end program t
