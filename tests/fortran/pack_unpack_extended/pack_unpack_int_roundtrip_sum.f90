! vybe-test: fortran/pack_unpack_extended/pack_unpack_int_roundtrip_sum
! origin: languages/fortran/tests/fortran/test_pack_unpack_extended.rs
program t
integer :: src(6) = [1, 2, 3, 4, 5, 6]
logical :: mask(6) = [.true., .false., .true., .false., .true., .false.]
integer :: tmp(3), dst(6)
integer :: fill(6) = [0, 0, 0, 0, 0, 0]
tmp = pack(src, mask)
dst = unpack(tmp, mask, fill)
if ((sum(dst)) /= 9) then
    print *, "FAIL: want [9] got [", sum(dst), "]"
    stop 1
end if
end program t
