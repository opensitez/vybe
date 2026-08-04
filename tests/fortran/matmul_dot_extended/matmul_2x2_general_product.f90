! vybe-test: fortran/matmul_dot_extended/matmul_2x2_general_product
! origin: languages/fortran/tests/fortran/test_matmul_dot_extended.rs
program t
integer :: a(2,2), b(2,2), c(2,2)
a(1,1)=1; a(1,2)=2; a(2,1)=3; a(2,2)=4
b(1,1)=5; b(1,2)=6; b(2,1)=7; b(2,2)=8
c = matmul(a, b)
if ((c(1,1)) /= 19) then
    print *, "FAIL: want [19] got [", c(1,1), "]"
    stop 1
end if
if ((c(1,2)) /= 22) then
    print *, "FAIL: want [22] got [", c(1,2), "]"
    stop 1
end if
if ((c(2,1)) /= 43) then
    print *, "FAIL: want [43] got [", c(2,1), "]"
    stop 1
end if
if ((c(2,2)) /= 50) then
    print *, "FAIL: want [50] got [", c(2,2), "]"
    stop 1
end if
end program t
