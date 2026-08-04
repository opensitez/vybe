! vybe-test: fortran/matmul_dot_extended/matmul_zero_matrix_2x2
! origin: languages/fortran/tests/fortran/test_matmul_dot_extended.rs
program t
integer :: z(2,2), b(2,2), c(2,2)
z = 0
b(1,1)=9; b(1,2)=-4; b(2,1)=3; b(2,2)=7
c = matmul(z, b)
if ((c(1,1)) /= 0) then
    print *, "FAIL: want [0] got [", c(1,1), "]"
    stop 1
end if
if ((c(1,2)) /= 0) then
    print *, "FAIL: want [0] got [", c(1,2), "]"
    stop 1
end if
if ((c(2,1)) /= 0) then
    print *, "FAIL: want [0] got [", c(2,1), "]"
    stop 1
end if
if ((c(2,2)) /= 0) then
    print *, "FAIL: want [0] got [", c(2,2), "]"
    stop 1
end if
end program t
