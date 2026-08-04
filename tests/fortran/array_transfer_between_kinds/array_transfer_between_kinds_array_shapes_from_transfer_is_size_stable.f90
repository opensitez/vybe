! vybe-test: fortran/array_transfer_between_kinds/array_transfer_between_kinds_array_shapes_from_transfer_is_size_stable
! origin: languages/fortran/tests/fortran/test_array_transfer_between_kinds.rs

program array_transfer_between_kinds_array_shapes_from_transfer_is_size_stable
    integer :: packed(2)
    integer :: flat(2)
    packed = (/7, 8/)
    flat = transfer(packed, flat)
    if ((size(packed)) /= 2) then
    print *, "FAIL: want [2] got [", size(packed), "]"
    stop 1
end if
    if ((size(flat)) /= 2) then
    print *, "FAIL: want [2] got [", size(flat), "]"
    stop 1
end if
    if ((flat(1)) /= 7) then
    print *, "FAIL: want [7] got [", flat(1), "]"
    stop 1
end if
    if ((flat(2)) /= 8) then
    print *, "FAIL: want [8] got [", flat(2), "]"
    stop 1
end if
end program array_transfer_between_kinds_array_shapes_from_transfer_is_size_stable
