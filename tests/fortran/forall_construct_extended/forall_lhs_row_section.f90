! vybe-test: fortran/forall_construct_extended/forall_lhs_row_section
! origin: languages/fortran/tests/fortran/test_forall_construct_extended.rs
program t
integer :: m(3,4)
m = 0
forall (i = 1:3)
m(i, 1:4) = i * 10
end forall
if ((m(1,1)) /= 10) then
    print *, "FAIL: want [10] got [", m(1,1), "]"
    stop 1
end if
if ((m(2,1)) /= 20) then
    print *, "FAIL: want [20] got [", m(2,1), "]"
    stop 1
end if
if ((m(3,4)) /= 30) then
    print *, "FAIL: want [30] got [", m(3,4), "]"
    stop 1
end if
end program t
