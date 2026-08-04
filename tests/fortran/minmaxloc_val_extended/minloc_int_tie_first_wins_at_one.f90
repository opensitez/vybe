! vybe-test: fortran/minmaxloc_val_extended/minloc_int_tie_first_wins_at_one
! origin: languages/fortran/tests/fortran/test_minmaxloc_val_extended.rs
program t
integer :: a(5) = [2, 2, 5, 2, 8]
integer :: loc(1)
loc = minloc(a)
if ((loc(1)) /= 1) then
    print *, "FAIL: want [1] got [", loc(1), "]"
    stop 1
end if
end program t
