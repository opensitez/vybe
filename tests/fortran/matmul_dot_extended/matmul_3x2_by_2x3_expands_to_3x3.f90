! vybe-test: fortran/matmul_dot_extended/matmul_3x2_by_2x3_expands_to_3x3
! origin: languages/fortran/tests/fortran/test_matmul_dot_extended.rs
program t
integer :: a(3,2), b(2,3), c(3,3)
a(1,1)=1; a(1,2)=2; a(2,1)=3; a(2,2)=4; a(3,1)=5; a(3,2)=6
b(1,1)=1; b(1,2)=0; b(1,3)=-1; b(2,1)=0; b(2,2)=1; b(2,3)=0
c = matmul(a, b)
if ((c(1,1)) /= 1) then
    print *, "FAIL: want [1] got [", c(1,1), "]"
    stop 1
end if
if ((c(1,3)) /= -1) then
    print *, "FAIL: want [-1] got [", c(1,3), "]"
    stop 1
end if
if ((c(3,1)) /= 5) then
    print *, "FAIL: want [5] got [", c(3,1), "]"
    stop 1
end if
if ((c(3,3)) /= -5) then
    print *, "FAIL: want [-5] got [", c(3,3), "]"
    stop 1
end if
end program t
