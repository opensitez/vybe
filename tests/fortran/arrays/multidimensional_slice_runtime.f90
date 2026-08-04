! vybe-test: fortran/arrays/multidimensional_slice_runtime
! origin: languages/fortran/tests/fortran/test_arrays.rs

program test
    integer :: a(4,4)
    integer :: b(2,2)

    a = 0
    a(2,2) = 11
    a(2,3) = 12
    a(3,2) = 21
    a(3,3) = 22

    b = a(2:3, 2:3)
    if ((b(1,1)) /= 11) then
    print *, "FAIL: want [11] got [", b(1,1), "]"
    stop 1
end if
    if ((b(1,2)) /= 12) then
    print *, "FAIL: want [12] got [", b(1,2), "]"
    stop 1
end if
    if ((b(2,1)) /= 21) then
    print *, "FAIL: want [21] got [", b(2,1), "]"
    stop 1
end if
    if ((b(2,2)) /= 22) then
    print *, "FAIL: want [22] got [", b(2,2), "]"
    stop 1
end if
end program test
