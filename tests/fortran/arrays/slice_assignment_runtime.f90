! vybe-test: fortran/arrays/slice_assignment_runtime
! origin: languages/fortran/tests/fortran/test_arrays.rs

program test
    integer :: a(5) = [1, 2, 3, 4, 5]

    a(2:4) = [20, 30, 40]

    if ((a(1)) /= 1) then
    print *, "FAIL: want [1] got [", a(1), "]"
    stop 1
end if
    if ((a(2)) /= 20) then
    print *, "FAIL: want [20] got [", a(2), "]"
    stop 1
end if
    if ((a(4)) /= 40) then
    print *, "FAIL: want [40] got [", a(4), "]"
    stop 1
end if
    if ((a(5)) /= 5) then
    print *, "FAIL: want [5] got [", a(5), "]"
    stop 1
end if
end program test
