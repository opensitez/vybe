! vybe-test: fortran/intrinsics_extended/not_basic
! origin: languages/fortran/tests/fortran/test_intrinsics_extended.rs
program t
integer :: x
x = not(0)
if ((x) /= -1) then
    print *, "FAIL: want [-1] got [", x, "]"
    stop 1
end if
end program t
