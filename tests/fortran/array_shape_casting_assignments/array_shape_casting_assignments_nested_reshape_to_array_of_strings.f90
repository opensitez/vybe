! vybe-test: fortran/array_shape_casting_assignments/array_shape_casting_assignments_nested_reshape_to_array_of_strings
! origin: languages/fortran/tests/fortran/test_array_shape_casting_assignments.rs

program t
    character(len=2) :: packed(2)
    integer :: flat(4)
    flat = (/11, 22, 33, 44/)
    packed = reshape(transfer(flat, (/''/)), (/2/))
    if (trim(packed(1)) /= "") then
    print *, "FAIL: want [] got [", packed(1), "]"
    stop 1
end if
    if (trim(packed(2)) /= "") then
    print *, "FAIL: want [] got [", packed(2), "]"
    stop 1
end if
    if ((len_trim(packed(1))) /= 0) then
    print *, "FAIL: want [0] got [", len_trim(packed(1)), "]"
    stop 1
end if
end program t
