! vybe-test: fortran/arrays/matmul_basic
! origin: languages/fortran/tests/fortran/test_arrays.rs

program test
    integer :: a(2,2), b(2,2), ident(2,2)
    integer :: c(2,2)
    a(1,1) = 1
    a(1,2) = 2
    a(2,1) = 3
    a(2,2) = 4
    ident = 0
    ident(1,1) = 1
    ident(2,2) = 1
    c = matmul(a, ident)
    if ((c(1,1)) /= 1) then
    print *, "FAIL: want [1] got [", c(1,1), "]"
    stop 1
end if
    if ((c(1,2)) /= 2) then
    print *, "FAIL: want [2] got [", c(1,2), "]"
    stop 1
end if
    if ((c(2,1)) /= 3) then
    print *, "FAIL: want [3] got [", c(2,1), "]"
    stop 1
end if
    if ((c(2,2)) /= 4) then
    print *, "FAIL: want [4] got [", c(2,2), "]"
    stop 1
end if
end program test
