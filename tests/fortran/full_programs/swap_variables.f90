! vybe-test: fortran/full_programs/swap_variables
! origin: languages/fortran/tests/fortran/test_full_programs.rs
program t
integer :: a, b, tmp
a = 10
b = 20
tmp = a
a = b
b = tmp
if ((a) /= 20) then
    print *, "FAIL: want [20] got [", a, "]"
    stop 1
end if
if ((b) /= 10) then
    print *, "FAIL: want [10] got [", b, "]"
    stop 1
end if
end program t
