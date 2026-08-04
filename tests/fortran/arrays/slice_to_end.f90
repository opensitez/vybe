! vybe-test: fortran/arrays/slice_to_end
! origin: languages/fortran/tests/fortran/test_arrays.rs

program test
    integer :: a(5) = [1, 2, 3, 4, 5]
    integer :: b(3)
    b = a(3:)
    if ((b(1)) /= 3) then
    print *, "FAIL: want [3] got [", b(1), "]"
    stop 1
end if
    if ((b(3)) /= 5) then
    print *, "FAIL: want [5] got [", b(3), "]"
    stop 1
end if
end program test
