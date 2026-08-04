! vybe-test: fortran/minmaxloc_val_extended/maxloc_int_tie_first_wins_at_two
! origin: languages/fortran/tests/fortran/test_minmaxloc_val_extended.rs
program t
integer :: a(5) = [4, 7, 7, 3, 7]
integer :: loc(1)
loc = maxloc(a)
if ((loc(1)) /= 2) then
    print *, "FAIL: want [2] got [", loc(1), "]"
    stop 1
end if
end program t
