! vybe-test: fortran/forall_construct_extended/forall_3d_fill_cube
! origin: languages/fortran/tests/fortran/test_forall_construct_extended.rs
program t
integer :: a(2,2,2)
a = 0
forall (i = 1:2, j = 1:2, k = 1:2)
a(i,j,k) = i * 100 + j * 10 + k
end forall
if ((a(1,1,1)) /= 111) then
    print *, "FAIL: want [111] got [", a(1,1,1), "]"
    stop 1
end if
if ((a(2,1,2)) /= 212) then
    print *, "FAIL: want [212] got [", a(2,1,2), "]"
    stop 1
end if
if ((a(2,2,2)) /= 222) then
    print *, "FAIL: want [222] got [", a(2,2,2), "]"
    stop 1
end if
end program t
