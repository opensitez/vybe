! vybe-test: fortran/array_dope_vector_copying/array_dope_vector_copying_section_copy_via_assumed_shape_argument
! origin: languages/fortran/tests/fortran/test_array_dope_vector_copying.rs

program t
    integer, allocatable :: source(:)
    integer, allocatable :: target(:)
    source = (/ 10, 20, 30, 40, 50 /)
    call copy_middle(source, 2, 4, target)
    if ((size(target)) /= 3) then
    print *, "FAIL: want [3] got [", size(target), "]"
    stop 1
end if
    if ((sum(target)) /= 90) then
    print *, "FAIL: want [90] got [", sum(target), "]"
    stop 1
end if
    if ((target(1)) /= 20) then
    print *, "FAIL: want [20] got [", target(1), "]"
    stop 1
end if
    if ((target(size(target))) /= 40) then
    print *, "FAIL: want [40] got [", target(size(target)), "]"
    stop 1
end if
contains
    subroutine copy_middle(values, i_start, i_end, out_values)
        integer, intent(in) :: values(:)
        integer, intent(in) :: i_start
        integer, intent(in) :: i_end
        integer, allocatable, intent(out) :: out_values(:)
        out_values = values(i_start:i_end)
    end subroutine copy_middle
end program t
