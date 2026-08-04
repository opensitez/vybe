! vybe-test: fortran/forall_construct_extended/forall_array_rhs_scaled_row
! origin: languages/fortran/tests/fortran/test_forall_construct_extended.rs
program t
integer :: u(4) = [1, 2, 3, 4]
integer :: m(4,4)
m = 0
forall (i = 1:4)
m(i, 1:4) = u(1:4) * i
end forall
if ((m(2,3)) /= 6) then
    print *, "FAIL: want [6] got [", m(2,3), "]"
    stop 1
end if
if ((m(4,1)) /= 4) then
    print *, "FAIL: want [4] got [", m(4,1), "]"
    stop 1
end if
if ((m(1,4)) /= 4) then
    print *, "FAIL: want [4] got [", m(1,4), "]"
    stop 1
end if
end program t
