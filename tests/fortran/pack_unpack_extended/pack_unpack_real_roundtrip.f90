! vybe-test: fortran/pack_unpack_extended/pack_unpack_real_roundtrip
! origin: languages/fortran/tests/fortran/test_pack_unpack_extended.rs
program t
real :: src(4) = [1.0, 2.0, 3.0, 4.0]
logical :: mask(4) = [.true., .false., .true., .false.]
real :: tmp(2), dst(4)
real :: fill(4) = [0.0, 0.0, 0.0, 0.0]
tmp = pack(src, mask)
dst = unpack(tmp, mask, fill)
if ((int(dst(1) + dst(3))) /= 4) then
    print *, "FAIL: want [4] got [", int(dst(1) + dst(3)), "]"
    stop 1
end if
end program t
