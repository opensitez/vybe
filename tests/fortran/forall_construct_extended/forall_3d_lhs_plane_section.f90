! vybe-test: fortran/forall_construct_extended/forall_3d_lhs_plane_section
! origin: languages/fortran/tests/fortran/test_forall_construct_extended.rs
program t
integer :: m(2,3,2)
m = 0
forall (k = 1:2, i = 1:2)
m(i, 1:3, k) = i * 10 + k
end forall
if ((m(1,2,1)) /= 11) then
    print *, "FAIL: want [11] got [", m(1,2,1), "]"
    stop 1
end if
if ((m(2,3,2)) /= 22) then
    print *, "FAIL: want [22] got [", m(2,3,2), "]"
    stop 1
end if
if ((m(1,1,2)) /= 12) then
    print *, "FAIL: want [12] got [", m(1,1,2), "]"
    stop 1
end if
end program t
