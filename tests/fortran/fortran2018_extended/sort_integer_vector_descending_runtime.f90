! vybe-test: fortran/fortran2018_extended/sort_integer_vector_descending_runtime
! origin: languages/fortran/tests/fortran/test_fortran2018_extended.rs

program t
    integer :: a(4) = [3, 1, 4, 2]
    call sort(a, reverse=.true.)
    if ((a(1)) /= 4) then
    print *, "FAIL: want [4] got [", a(1), "]"
    stop 1
end if
    if ((a(2)) /= 3) then
    print *, "FAIL: want [3] got [", a(2), "]"
    stop 1
end if
    if ((a(3)) /= 2) then
    print *, "FAIL: want [2] got [", a(3), "]"
    stop 1
end if
    if ((a(4)) /= 1) then
    print *, "FAIL: want [1] got [", a(4), "]"
    stop 1
end if
end program t
