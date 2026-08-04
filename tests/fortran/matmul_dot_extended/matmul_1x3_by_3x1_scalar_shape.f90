! vybe-test: fortran/matmul_dot_extended/matmul_1x3_by_3x1_scalar_shape
! origin: languages/fortran/tests/fortran/test_matmul_dot_extended.rs
program t
integer :: a(1,3), b(3,1), c(1,1)
a(1,1)=1; a(1,2)=2; a(1,3)=3
b(1,1)=4; b(2,1)=5; b(3,1)=6
c = matmul(a, b)
if ((c(1,1)) /= 32) then
    print *, "FAIL: want [32] got [", c(1,1), "]"
    stop 1
end if
end program t
