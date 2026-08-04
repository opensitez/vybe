! vybe-test: fortran/forall_construct_extended/forall_mask_i_plus_j_equals
! origin: languages/fortran/tests/fortran/test_forall_construct_extended.rs
program t
integer :: m(4,4)
m = 0
forall (i = 1:4, j = 1:4, i + j == 5)
m(i,j) = i * j
end forall
if ((m(1,4)) /= 4) then
    print *, "FAIL: want [4] got [", m(1,4), "]"
    stop 1
end if
if ((m(2,3)) /= 6) then
    print *, "FAIL: want [6] got [", m(2,3), "]"
    stop 1
end if
if ((m(1,1)) /= 0) then
    print *, "FAIL: want [0] got [", m(1,1), "]"
    stop 1
end if
end program t
