! vybe-test: fortran/array_elemental_lowering/array_elemental_lowering_elemental_reshape_replay
! origin: languages/fortran/tests/fortran/test_array_elemental_lowering.rs

program array_elemental_lowering_elemental_reshape_replay
    integer, allocatable :: flat(:)
    integer, allocatable :: matrix(:,:)
    integer :: corner
    flat = (/ 1, 2, 3, 4, 5, 6 /)
    matrix = reshape(abs(flat), (/2, 3/))
    corner = matrix(1,1) + matrix(2,3)
    if ((size(matrix,1)) /= 2) then
    print *, "FAIL: want [2] got [", size(matrix,1), "]"
    stop 1
end if
    if ((size(matrix,2)) /= 3) then
    print *, "FAIL: want [3] got [", size(matrix,2), "]"
    stop 1
end if
    if ((corner) /= 7) then
    print *, "FAIL: want [7] got [", corner, "]"
    stop 1
end if
    if ((sum(matrix)) /= 21) then
    print *, "FAIL: want [21] got [", sum(matrix), "]"
    stop 1
end if
end program array_elemental_lowering_elemental_reshape_replay
