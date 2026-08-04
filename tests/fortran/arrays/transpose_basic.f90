! vybe-test: fortran/arrays/transpose_basic
! origin: languages/fortran/tests/fortran/test_arrays.rs

program test
    integer :: a(2,3)
    integer :: b(3,2)
    a(1,1) = 1
    a(1,2) = 2
    a(1,3) = 3
    a(2,1) = 4
    a(2,2) = 5
    a(2,3) = 6
    b = transpose(a)
    if ((b(1,1)) /= 1) then
    print *, "FAIL: want [1] got [", b(1,1), "]"
    stop 1
end if
    if ((b(3,2)) /= 6) then
    print *, "FAIL: want [6] got [", b(3,2), "]"
    stop 1
end if
end program test
