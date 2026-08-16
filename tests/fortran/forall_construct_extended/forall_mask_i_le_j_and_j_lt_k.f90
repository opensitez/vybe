! vybe-test: fortran/forall_construct_extended/forall_mask_i_le_j_and_j_lt_k
! origin: languages/fortran/tests/fortran/test_forall_construct_extended.rs
program t
integer :: a(3,3,3)
a = 0
forall (i = 1:3, j = 1:3, k = 1:3, i <= j .and. j < k)
a(i,j,k) = 10 * i + j + k
end forall
if ((a(1,1,2)) /= 13) then
    print *, "FAIL: want [13] got [", a(1,1,2), "]"
    stop 1
end if
if ((a(1,2,3)) /= 15) then
    print *, "FAIL: want [15] got [", a(1,2,3), "]"
    stop 1
end if
if ((a(2,2,2)) /= 0) then
    print *, "FAIL: want [0] got [", a(2,2,2), "]"
    stop 1
end if
end program t
