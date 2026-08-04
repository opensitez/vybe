! vybe-test: fortran/forall_construct_extended/forall_mask_i_times_j_gt
! origin: languages/fortran/tests/fortran/test_forall_construct_extended.rs
program t
integer :: m(4,4)
m = 0
forall (i = 1:4, j = 1:4, i * j > 6)
m(i,j) = i + j
end forall
if ((m(2,4)) /= 6) then
    print *, "FAIL: want [6] got [", m(2,4), "]"
    stop 1
end if
if ((m(3,3)) /= 6) then
    print *, "FAIL: want [6] got [", m(3,3), "]"
    stop 1
end if
if ((m(2,2)) /= 0) then
    print *, "FAIL: want [0] got [", m(2,2), "]"
    stop 1
end if
end program t
