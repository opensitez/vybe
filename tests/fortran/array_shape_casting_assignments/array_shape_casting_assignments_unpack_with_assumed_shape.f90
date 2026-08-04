! vybe-test: fortran/array_shape_casting_assignments/array_shape_casting_assignments_unpack_with_assumed_shape
! origin: languages/fortran/tests/fortran/test_array_shape_casting_assignments.rs

program array_shape_casting_assignments_unpack_with_assumed_shape
    integer :: a(2, 2)
    integer :: b(4)
    call write_back(a, b)
    if ((b(1)) /= 1) then
    print *, "FAIL: want [1] got [", b(1), "]"
    stop 1
end if
    if ((b(2)) /= 2) then
    print *, "FAIL: want [2] got [", b(2), "]"
    stop 1
end if
    if ((b(3)) /= 3) then
    print *, "FAIL: want [3] got [", b(3), "]"
    stop 1
end if
    if ((b(4)) /= 4) then
    print *, "FAIL: want [4] got [", b(4), "]"
    stop 1
end if
contains
    subroutine write_back(src, dst)
        integer, intent(in)  :: src(:, :)
        integer, intent(out) :: dst(:)
        dst = reshape(src, (/4/))
    end subroutine write_back
end program array_shape_casting_assignments_unpack_with_assumed_shape
