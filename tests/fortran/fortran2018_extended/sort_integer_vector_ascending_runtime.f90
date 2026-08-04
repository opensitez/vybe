! vybe-test: fortran/fortran2018_extended/sort_integer_vector_ascending_runtime
! origin: languages/fortran/tests/fortran/test_fortran2018_extended.rs

program t
    integer :: a(5) = [3, 1, 4, 1, 5]
    call sort(a)
    if ((a(1)) /= 1) then
    print *, "FAIL: want [1] got [", a(1), "]"
    stop 1
end if
    if ((a(2)) /= 1) then
    print *, "FAIL: want [1] got [", a(2), "]"
    stop 1
end if
    if ((a(3)) /= 3) then
    print *, "FAIL: want [3] got [", a(3), "]"
    stop 1
end if
    if ((a(4)) /= 4) then
    print *, "FAIL: want [4] got [", a(4), "]"
    stop 1
end if
    if ((a(5)) /= 5) then
    print *, "FAIL: want [5] got [", a(5), "]"
    stop 1
end if
end program t
