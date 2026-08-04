! vybe-test: fortran/coarray_sync_extended/event_wait_until_count_one
! origin: languages/fortran/tests/fortran/test_coarray_sync_extended.rs
program t
use iso_fortran_env
type(event_type) :: ev[*]
event post(ev)
event wait(ev, until_count=1)
if (trim('ready') /= "ready") then
    print *, "FAIL: want [ready] got [", 'ready', "]"
    stop 1
end if
end program t
