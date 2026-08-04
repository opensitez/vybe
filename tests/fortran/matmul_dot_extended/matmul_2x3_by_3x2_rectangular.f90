! vybe-test: fortran/matmul_dot_extended/matmul_2x3_by_3x2_rectangular
! origin: languages/fortran/tests/fortran/test_matmul_dot_extended.rs
program t
integer :: a(2,3), b(3,2), c(2,2)
a(1,1)=1; a(1,2)=2; a(1,3)=3; a(2,1)=4; a(2,2)=5; a(2,3)=6
b(1,1)=7; b(1,2)=1; b(2,1)=8; b(2,2)=0; b(3,1)=9; b(3,2)=-1
c = matmul(a, b)
if ((c(1,1)) /= 50) then
    print *, "FAIL: want [50] got [", c(1,1), "]"
    stop 1
end if
if ((c(1,2)) /= -2) then
    print *, "FAIL: want [-2] got [", c(1,2), "]"
    stop 1
end if
if ((c(2,1)) /= 122) then
    print *, "FAIL: want [122] got [", c(2,1), "]"
    stop 1
end if
if ((c(2,2)) /= -2) then
    print *, "FAIL: want [-2] got [", c(2,2), "]"
    stop 1
end if
end program t
