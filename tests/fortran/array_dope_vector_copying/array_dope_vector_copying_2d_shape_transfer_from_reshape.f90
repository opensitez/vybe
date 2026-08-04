! vybe-test: fortran/array_dope_vector_copying/array_dope_vector_copying_2d_shape_transfer_from_reshape
! origin: languages/fortran/tests/fortran/test_array_dope_vector_copying.rs

program array_dope_vector_copying_2d_shape_transfer_from_reshape
    integer, allocatable :: source(:)
    integer, allocatable :: target(:,:)
    integer :: t
    source = (/ 1, 2, 3, 4, 5, 6 /)
    target = reshape(source, (/2,3/))
    t = target(2, 3)
    if ((size(target,1)) /= 2) then
    print *, "FAIL: want [2] got [", size(target,1), "]"
    stop 1
end if
    if ((size(target,2)) /= 3) then
    print *, "FAIL: want [3] got [", size(target,2), "]"
    stop 1
end if
    if ((sum(target)) /= 21) then
    print *, "FAIL: want [21] got [", sum(target), "]"
    stop 1
end if
    if ((t) /= 6) then
    print *, "FAIL: want [6] got [", t, "]"
    stop 1
end if
end program array_dope_vector_copying_2d_shape_transfer_from_reshape
