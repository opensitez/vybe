! vybe-test: fortran/fortran2018_extended/reduce_custom_min_procedure
! origin: languages/fortran/tests/fortran/test_fortran2018_extended.rs
program t
integer :: a(5) = [8, 3, 9, 1, 6]
if ((reduce(a, pick_min)) /= 1) then
    print *, "FAIL: want [1] got [", reduce(a, pick_min), "]"
    stop 1
end if
contains
pure function pick_min(x, y) result(r)
integer, intent(in) :: x, y
integer :: r
r = min(x, y)
end function pick_min
end program t
