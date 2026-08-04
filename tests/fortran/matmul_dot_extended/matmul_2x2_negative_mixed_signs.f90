! vybe-test: fortran/matmul_dot_extended/matmul_2x2_negative_mixed_signs
! origin: languages/fortran/tests/fortran/test_matmul_dot_extended.rs
program t
integer :: a(2,2), b(2,2), c(2,2)
a(1,1)=1; a(1,2)=-1; a(2,1)=2; a(2,2)=-2
b(1,1)=3; b(1,2)=4; b(2,1)=-1; b(2,2)=0
c = matmul(a, b)
if ((c(1,1)) /= 4) then
    print *, "FAIL: want [4] got [", c(1,1), "]"
    stop 1
end if
if ((c(1,2)) /= 4) then
    print *, "FAIL: want [4] got [", c(1,2), "]"
    stop 1
end if
if ((c(2,1)) /= 8) then
    print *, "FAIL: want [8] got [", c(2,1), "]"
    stop 1
end if
if ((c(2,2)) /= 8) then
    print *, "FAIL: want [8] got [", c(2,2), "]"
    stop 1
end if
end program t
