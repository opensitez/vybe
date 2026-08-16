! vybe-test: fortran/pack_unpack_extended/pack_int_vector_alternating_mask_six_slots
! origin: languages/fortran/tests/fortran/test_pack_unpack_extended.rs
program t
integer :: a(6) = [1, 2, 3, 4, 5, 6]
logical :: mask(6) = [.true., .false., .true., .false., .true., .false.]
integer :: vec(6) = [99, 99, 99, 99, 99, 99]
integer :: b(6)
b = pack(a, mask, vec)
if ((b(1)) /= 1) then
    print *, "FAIL: want [1] got [", b(1), "]"
    stop 1
end if
if ((b(2)) /= 3) then
    print *, "FAIL: want [3] got [", b(2), "]"
    stop 1
end if
if ((b(3)) /= 5) then
    print *, "FAIL: want [5] got [", b(3), "]"
    stop 1
end if
if ((b(6)) /= 99) then
    print *, "FAIL: want [99] got [", b(6), "]"
    stop 1
end if
end program t
