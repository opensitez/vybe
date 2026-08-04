! vybe-test: fortran/arrays/array_element_write
! origin: languages/fortran/tests/fortran/test_arrays.rs

program test
    integer :: a(3)
    a(1) = 10
    a(2) = 20
    a(3) = 30
    if ((a(1)) /= 10) then
    print *, "FAIL: want [10] got [", a(1), "]"
    stop 1
end if
    if ((a(2)) /= 20) then
    print *, "FAIL: want [20] got [", a(2), "]"
    stop 1
end if
    if ((a(3)) /= 30) then
    print *, "FAIL: want [30] got [", a(3), "]"
    stop 1
end if
end program test
