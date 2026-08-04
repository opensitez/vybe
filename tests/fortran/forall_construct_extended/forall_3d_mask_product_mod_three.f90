! vybe-test: fortran/forall_construct_extended/forall_3d_mask_product_mod_three
! origin: languages/fortran/tests/fortran/test_forall_construct_extended.rs
program t
integer :: a(3,3,3)
a = 0
forall (i = 1:3, j = 1:3, k = 1:3, mod(i * j * k, 3) == 0)
a(i,j,k) = i + j + k
end forall
if ((a(1,1,3)) /= 5) then
    print *, "FAIL: want [5] got [", a(1,1,3), "]"
    stop 1
end if
if ((a(2,3,1)) /= 6) then
    print *, "FAIL: want [6] got [", a(2,3,1), "]"
    stop 1
end if
if ((a(2,2,2)) /= 0) then
    print *, "FAIL: want [0] got [", a(2,2,2), "]"
    stop 1
end if
end program t
