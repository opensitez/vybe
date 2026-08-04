! vybe-test: fortran/matmul_dot_extended/matmul_row_vector_times_matrix
! origin: languages/fortran/tests/fortran/test_matmul_dot_extended.rs
program t
integer :: v(3), b(3,2), c(2)
v = [1, 2, 3]
b(1,1)=1; b(1,2)=0; b(2,1)=0; b(2,2)=1; b(3,1)=1; b(3,2)=1
c = matmul(v, b)
if ((c(1)) /= 4) then
    print *, "FAIL: want [4] got [", c(1), "]"
    stop 1
end if
if ((c(2)) /= 5) then
    print *, "FAIL: want [5] got [", c(2), "]"
    stop 1
end if
end program t
