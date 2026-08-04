! vybe-test: fortran/pack_unpack_extended/pack_int_vector_single_pad_slot
! origin: languages/fortran/tests/fortran/test_pack_unpack_extended.rs
program t
integer :: a(3) = [10, 20, 30]
logical :: mask(3) = [.true., .false., .false.]
integer :: vec(2) = [99, 99]
integer :: b(2)
b = pack(a, mask, vec)
if ((b(1)) /= 10) then
    print *, "FAIL: want [10] got [", b(1), "]"
    stop 1
end if
if ((b(2)) /= 99) then
    print *, "FAIL: want [99] got [", b(2), "]"
    stop 1
end if
end program t
