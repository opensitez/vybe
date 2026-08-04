! vybe-test: fortran/forall_construct_extended/forall_3d_stride_subcube
! origin: languages/fortran/tests/fortran/test_forall_construct_extended.rs
program t
integer :: a(4,4,4)
a = 0
forall (i = 1:4:2, j = 1:4:2, k = 1:4:2)
a(i,j,k) = i * j * k
end forall
if ((a(1,1,1)) /= 1) then
    print *, "FAIL: want [1] got [", a(1,1,1), "]"
    stop 1
end if
if ((a(3,3,3)) /= 27) then
    print *, "FAIL: want [27] got [", a(3,3,3), "]"
    stop 1
end if
if ((a(2,2,2)) /= 0) then
    print *, "FAIL: want [0] got [", a(2,2,2), "]"
    stop 1
end if
end program t
