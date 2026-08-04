! vybe-test: fortran/coarray_sync_extended/sync_team_on_initial_team
! origin: languages/fortran/tests/fortran/test_coarray_sync_extended.rs
program t
use iso_fortran_env
type(team_type) :: init
type(event_type) :: ev[*]
init = get_team(initial_team)
sync team (init)
if ((team_number(init)) /= -1) then
    print *, "FAIL: want [-1] got [", team_number(init), "]"
    stop 1
end if
end program t
