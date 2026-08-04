! vybe-test: fortran/pack_unpack_extended/pack_unpack_int_identity_all_true
! origin: languages/fortran/tests/fortran/test_pack_unpack_extended.rs
program t
integer :: src(3) = [7, 8, 9]
logical :: mask(3) = [.true., .true., .true.]
integer :: tmp(3), dst(3)
integer :: fill(3) = [0, 0, 0]
tmp = pack(src, mask)
dst = unpack(tmp, mask, fill)
if ((dst(1)) /= 7) then
    print *, "FAIL: want [7] got [", dst(1), "]"
    stop 1
end if
if ((dst(2)) /= 8) then
    print *, "FAIL: want [8] got [", dst(2), "]"
    stop 1
end if
if ((dst(3)) /= 9) then
    print *, "FAIL: want [9] got [", dst(3), "]"
    stop 1
end if
end program t
