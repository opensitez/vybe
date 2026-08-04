! vybe-test: fortran/forall_construct_extended/forall_3d_mask_ordered_indices
! origin: languages/fortran/tests/fortran/test_forall_construct_extended.rs
program t
integer :: a(3,3,3)
a = 0
forall (i = 1:3, j = 1:3, k = 1:3, i <= j .and. j <= k)
a(i,j,k) = 1
end forall
if ((a(1,1,1)) /= 1) then
    print *, "FAIL: want [1] got [", a(1,1,1), "]"
    stop 1
end if
if ((a(1,2,3)) /= 1) then
    print *, "FAIL: want [1] got [", a(1,2,3), "]"
    stop 1
end if
if ((a(2,1,3)) /= 0) then
    print *, "FAIL: want [0] got [", a(2,1,3), "]"
    stop 1
end if
end program t
