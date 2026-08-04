! vybe-test: fortran/array_shape_casting_assignments/array_shape_casting_assignments_pack_unpack_pairings
! origin: languages/fortran/tests/fortran/test_array_shape_casting_assignments.rs

program array_shape_casting_assignments_pack_unpack_pairings
    integer :: src(2, 3)
    integer :: dst(3)
    integer :: packed(2)
    src = reshape((/1, 2, 3, 4, 5, 6/), (/2, 3/))
    packed = reshape(reshape(src, (/6/))(1:2), (/2/))
    dst = (/packed(1), packed(2), 9/)
    if ((dst(1)) /= 1) then
    print *, "FAIL: want [1] got [", dst(1), "]"
    stop 1
end if
    if ((dst(2)) /= 2) then
    print *, "FAIL: want [2] got [", dst(2), "]"
    stop 1
end if
    if ((dst(3)) /= 9) then
    print *, "FAIL: want [9] got [", dst(3), "]"
    stop 1
end if
    if ((sum(dst)) /= 12) then
    print *, "FAIL: want [12] got [", sum(dst), "]"
    stop 1
end if
end program array_shape_casting_assignments_pack_unpack_pairings
