! vybe-test: fortran/forall_construct_extended/forall_array_rhs_row_to_matrix
! origin: languages/fortran/tests/fortran/test_forall_construct_extended.rs
program t
integer :: src(4) = [5, 6, 7, 8]
integer :: m(3,4)
m = 0
forall (i = 1:3)
m(i, 1:4) = src(1:4)
end forall
if ((m(1,3)) /= 7) then
    print *, "FAIL: want [7] got [", m(1,3), "]"
    stop 1
end if
if ((m(2,4)) /= 8) then
    print *, "FAIL: want [8] got [", m(2,4), "]"
    stop 1
end if
if ((m(3,1)) /= 5) then
    print *, "FAIL: want [5] got [", m(3,1), "]"
    stop 1
end if
end program t
