! vybe-test: fortran/matmul_dot_extended/matmul_matrix_times_column_vector
! origin: languages/fortran/tests/fortran/test_matmul_dot_extended.rs
program t
integer :: a(2,3), v(3), c(2)
a(1,1)=1; a(1,2)=2; a(1,3)=3; a(2,1)=4; a(2,2)=5; a(2,3)=6
v = [1, 0, -1]
c = matmul(a, v)
if ((c(1)) /= -2) then
    print *, "FAIL: want [-2] got [", c(1), "]"
    stop 1
end if
if ((c(2)) /= -2) then
    print *, "FAIL: want [-2] got [", c(2), "]"
    stop 1
end if
end program t
