! vybe-test: fortran/matmul_dot_extended/matmul_identity_3x3_preserves
! origin: languages/fortran/tests/fortran/test_matmul_dot_extended.rs
program t
integer :: ident(3,3), a(3,3), c(3,3)
ident = 0; ident(1,1)=1; ident(2,2)=1; ident(3,3)=1
a(1,1)=2; a(1,2)=3; a(1,3)=4; a(2,1)=5; a(2,2)=6; a(2,3)=7; a(3,1)=8; a(3,2)=9; a(3,3)=10
c = matmul(a, ident)
if ((c(2,2)) /= 6) then
    print *, "FAIL: want [6] got [", c(2,2), "]"
    stop 1
end if
if ((c(3,1)) /= 8) then
    print *, "FAIL: want [8] got [", c(3,1), "]"
    stop 1
end if
if ((sum(c)) /= 54) then
    print *, "FAIL: want [54] got [", sum(c), "]"
    stop 1
end if
end program t
