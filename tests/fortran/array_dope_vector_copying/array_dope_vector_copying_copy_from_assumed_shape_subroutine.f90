! vybe-test: fortran/array_dope_vector_copying/array_dope_vector_copying_copy_from_assumed_shape_subroutine
! origin: languages/fortran/tests/fortran/test_array_dope_vector_copying.rs

program array_dope_vector_copying_copy_from_assumed_shape_subroutine
    integer, allocatable :: source(:)
    integer, allocatable :: target(:)
    source = (/ 4, 3, 2, 1 /)
    call copy_out(source, target)
    if ((size(target)) /= 4) then
    print *, "FAIL: want [4] got [", size(target), "]"
    stop 1
end if
    if ((sum(target)) /= 10) then
    print *, "FAIL: want [10] got [", sum(target), "]"
    stop 1
end if
    if ((target(1)) /= 4) then
    print *, "FAIL: want [4] got [", target(1), "]"
    stop 1
end if
    if ((target(size(target))) /= 1) then
    print *, "FAIL: want [1] got [", target(size(target)), "]"
    stop 1
end if
contains
    subroutine copy_out(values, out_values)
        integer, intent(in) :: values(:)
        integer, allocatable, intent(out) :: out_values(:)
        out_values = values
    end subroutine copy_out
end program array_dope_vector_copying_copy_from_assumed_shape_subroutine
