! vybe-test: fortran/arrays/allocatable_assignment_from_array_function_result
! origin: languages/fortran/tests/fortran/test_arrays.rs

program test
    real, allocatable :: v(:)
    allocate(v(3))
    v = values()
    if ((v(1)) /= 1) then
    print *, "FAIL: want [1] got [", v(1), "]"
    stop 1
end if
    if ((v(2)) /= 2) then
    print *, "FAIL: want [2] got [", v(2), "]"
    stop 1
end if
    if ((v(3)) /= 3) then
    print *, "FAIL: want [3] got [", v(3), "]"
    stop 1
end if
contains
    pure function values() result(a)
        real :: a(3)
        a(1) = 1.0
        a(2) = 2.0
        a(3) = 3.0
    end function values
end program test
