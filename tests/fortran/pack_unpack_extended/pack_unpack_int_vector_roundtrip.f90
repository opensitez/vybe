! vybe-test: fortran/pack_unpack_extended/pack_unpack_int_vector_roundtrip
! origin: languages/fortran/tests/fortran/test_pack_unpack_extended.rs
program t
integer :: src(4) = [1, 3, 5, 7]
logical :: mask(4) = [.true., .false., .true., .false.]
integer :: vec(4) = [0, 0, 0, 0]
integer :: tmp(4), dst(4)
integer :: fill(4) = [9, 9, 9, 9]
tmp = pack(src, mask, vec)
dst = unpack(tmp(1:2), mask, fill)
if ((dst(1)) /= 1) then
    print *, "FAIL: want [1] got [", dst(1), "]"
    stop 1
end if
if ((dst(3)) /= 5) then
    print *, "FAIL: want [5] got [", dst(3), "]"
    stop 1
end if
end program t
