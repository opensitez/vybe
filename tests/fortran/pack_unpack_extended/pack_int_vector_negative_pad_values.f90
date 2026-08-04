! vybe-test: fortran/pack_unpack_extended/pack_int_vector_negative_pad_values
! origin: languages/fortran/tests/fortran/test_pack_unpack_extended.rs
program t
integer :: a(4) = [4, 8, 12, 16]
logical :: mask(4) = [.true., .true., .false., .false.]
integer :: vec(3) = [-5, -5, -5]
integer :: b(3)
b = pack(a, mask, vec)
if ((b(1)) /= 4) then
    print *, "FAIL: want [4] got [", b(1), "]"
    stop 1
end if
if ((b(2)) /= 8) then
    print *, "FAIL: want [8] got [", b(2), "]"
    stop 1
end if
if ((b(3)) /= -5) then
    print *, "FAIL: want [-5] got [", b(3), "]"
    stop 1
end if
end program t
