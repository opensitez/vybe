! vybe-test: fortran/associate_construct_extended/associate_array_2d_diagonal
! origin: languages/fortran/tests/fortran/test_associate_construct_extended.rs
program t
integer :: m(3,3)
m = 0
m(1,1) = 7
m(2,2) = 8
m(3,3) = 9
associate (d => m(2,2))
if ((d) /= 8) then
    print *, "FAIL: want [8] got [", d, "]"
    stop 1
end if
end associate
end program t
