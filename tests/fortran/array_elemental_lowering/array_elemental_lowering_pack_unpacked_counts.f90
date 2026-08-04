! vybe-test: fortran/array_elemental_lowering/array_elemental_lowering_pack_unpacked_counts
! origin: languages/fortran/tests/fortran/test_array_elemental_lowering.rs

program array_elemental_lowering_pack_unpacked_counts
    integer, allocatable :: values(:)
    integer :: packed_count
    integer :: unpack_count
    values = (/ 5, 0, -2, 8, 0, 1 /)
    packed_count = size(pack(values, values > 0))
    unpack_count = size(pack(values, values == 0))
    if ((packed_count) /= 3) then
    print *, "FAIL: want [3] got [", packed_count, "]"
    stop 1
end if
    if ((unpack_count) /= 2) then
    print *, "FAIL: want [2] got [", unpack_count, "]"
    stop 1
end if
    if ((values(1) + values(6)) /= 6) then
    print *, "FAIL: want [6] got [", values(1) + values(6), "]"
    stop 1
end if
end program array_elemental_lowering_pack_unpacked_counts
