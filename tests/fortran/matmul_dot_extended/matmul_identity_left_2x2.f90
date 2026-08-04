! vybe-test: fortran/matmul_dot_extended/matmul_identity_left_2x2
! origin: languages/fortran/tests/fortran/test_matmul_dot_extended.rs
program t
integer :: ident(2,2), a(2,2), c(2,2)
ident = 0; ident(1,1)=1; ident(2,2)=1
a(1,1)=7; a(1,2)=-3; a(2,1)=5; a(2,2)=2
c = matmul(ident, a)
if ((c(1,1)) /= 7) then
    print *, "FAIL: want [7] got [", c(1,1), "]"
    stop 1
end if
if ((c(1,2)) /= -3) then
    print *, "FAIL: want [-3] got [", c(1,2), "]"
    stop 1
end if
if ((c(2,1)) /= 5) then
    print *, "FAIL: want [5] got [", c(2,1), "]"
    stop 1
end if
if ((c(2,2)) /= 2) then
    print *, "FAIL: want [2] got [", c(2,2), "]"
    stop 1
end if
end program t
