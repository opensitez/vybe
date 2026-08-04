! vybe-test: fortran/array_shape_casting_assignments/array_shape_casting_assignments_transpose_view_shape
! origin: languages/fortran/tests/fortran/test_array_shape_casting_assignments.rs

program array_shape_casting_assignments_transpose_view_shape
    integer :: a(2, 3)
    integer :: b(3, 2)
    a = reshape((/1, 2, 3, 4, 5, 6/), (/2, 3/))
    b = transpose(a)
    if ((b(1, 1)) /= 1) then
    print *, "FAIL: want [1] got [", b(1, 1), "]"
    stop 1
end if
    if ((b(3, 2)) /= 6) then
    print *, "FAIL: want [6] got [", b(3, 2), "]"
    stop 1
end if
    if ((sum(b)) /= 21) then
    print *, "FAIL: want [21] got [", sum(b), "]"
    stop 1
end if
end program array_shape_casting_assignments_transpose_view_shape
