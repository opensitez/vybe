! vybe-test: fortran/forall_construct_extended/forall_array_rhs_col_broadcast
! origin: languages/fortran/tests/fortran/test_forall_construct_extended.rs
program t
integer :: v(3) = [1, 2, 3]
integer :: m(3,3)
m = 0
forall (j = 1:3)
m(1:3, j) = v(1:3)
end forall
if ((m(2,2)) /= 2) then
    print *, "FAIL: want [2] got [", m(2,2), "]"
    stop 1
end if
if ((m(3,1)) /= 3) then
    print *, "FAIL: want [3] got [", m(3,1), "]"
    stop 1
end if
if ((m(1,3)) /= 1) then
    print *, "FAIL: want [1] got [", m(1,3), "]"
    stop 1
end if
end program t
