! vybe-test: fortran/pack_unpack_extended/pack_unpack_2d_roundtrip_corner
! origin: languages/fortran/tests/fortran/test_pack_unpack_extended.rs
program t
integer :: src(2,2) = reshape([5, 6, 7, 8], [2, 2])
logical :: mask(2,2) = reshape([.true., .false., .false., .true.], [2, 2])
integer :: tmp(2), dst(2,2)
integer :: fill(2,2) = 0
tmp = pack(src, mask)
dst = unpack(tmp, mask, fill)
if ((dst(1,1)) /= 5) then
    print *, "FAIL: want [5] got [", dst(1,1), "]"
    stop 1
end if
if ((dst(2,2)) /= 8) then
    print *, "FAIL: want [8] got [", dst(2,2), "]"
    stop 1
end if
end program t
