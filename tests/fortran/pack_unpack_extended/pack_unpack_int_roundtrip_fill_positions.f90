! vybe-test: fortran/pack_unpack_extended/pack_unpack_int_roundtrip_fill_positions
! origin: languages/fortran/tests/fortran/test_pack_unpack_extended.rs
program t
integer :: src(4) = [2, 4, 6, 8]
logical :: mask(4) = [.true., .true., .false., .false.]
integer :: tmp(2), dst(4)
integer :: fill(4) = [100, 100, 100, 100]
tmp = pack(src, mask)
dst = unpack(tmp, mask, fill)
if ((dst(1)) /= 2) then
    print *, "FAIL: want [2] got [", dst(1), "]"
    stop 1
end if
if ((dst(2)) /= 4) then
    print *, "FAIL: want [4] got [", dst(2), "]"
    stop 1
end if
if ((dst(3)) /= 100) then
    print *, "FAIL: want [100] got [", dst(3), "]"
    stop 1
end if
end program t
