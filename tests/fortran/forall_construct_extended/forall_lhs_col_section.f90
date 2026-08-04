! vybe-test: fortran/forall_construct_extended/forall_lhs_col_section
! origin: languages/fortran/tests/fortran/test_forall_construct_extended.rs
program t
integer :: m(4,3)
m = 0
forall (j = 1:3)
m(1:4, j) = j
end forall
if ((m(1,2)) /= 2) then
    print *, "FAIL: want [2] got [", m(1,2), "]"
    stop 1
end if
if ((m(4,2)) /= 2) then
    print *, "FAIL: want [2] got [", m(4,2), "]"
    stop 1
end if
if ((m(2,1)) /= 1) then
    print *, "FAIL: want [1] got [", m(2,1), "]"
    stop 1
end if
end program t
