! vybe-test: fortran/coarray_sync_extended/event_post_then_query_count
! origin: languages/fortran/tests/fortran/test_coarray_sync_extended.rs
program t
use iso_fortran_env
type(event_type) :: ev[*]
integer :: count
event post(ev)
event query(ev, count)
if ((count) /= 1) then
    print *, "FAIL: want [1] got [", count, "]"
    stop 1
end if
end program t
