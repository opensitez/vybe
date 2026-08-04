! vybe-test: fortran/array_shape_casting_assignments/array_shape_casting_assignments_shape_function_stability
! origin: languages/fortran/tests/fortran/test_array_shape_casting_assignments.rs

program array_shape_casting_assignments_shape_function_stability
    integer :: source(2, 3)
    integer :: first(6)
    integer :: second(3, 2)
    source = reshape((/1, 2, 3, 4, 5, 6/), (/2, 3/))
    first = reshape(source, (/6/))
    second = reshape(first, (/3, 2/))
    if ((shape(first)(1)) /= 6) then
    print *, "FAIL: want [6] got [", shape(first)(1), "]"
    stop 1
end if
    if ((shape(second)(1)) /= 3) then
    print *, "FAIL: want [3] got [", shape(second)(1), "]"
    stop 1
end if
    if ((shape(second)(2)) /= 2) then
    print *, "FAIL: want [2] got [", shape(second)(2), "]"
    stop 1
end if
    if ((sum(second)) /= 21) then
    print *, "FAIL: want [21] got [", sum(second), "]"
    stop 1
end if
end program array_shape_casting_assignments_shape_function_stability
