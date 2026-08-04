! vybe-test: fortran/pack_unpack_extended/pack_unpack_char_roundtrip
! origin: languages/fortran/tests/fortran/test_pack_unpack_extended.rs
program t
character(len=1) :: src(4) = ['W', 'X', 'Y', 'Z']
logical :: mask(4) = [.true., .false., .true., .false.]
character(len=1) :: tmp(2), dst(4)
character(len=1) :: fill(4) = ['?', '?', '?', '?']
tmp = pack(src, mask)
dst = unpack(tmp, mask, fill)
if (trim(dst(1)) /= "W") then
    print *, "FAIL: want [W] got [", dst(1), "]"
    stop 1
end if
if (trim(dst(3)) /= "Y") then
    print *, "FAIL: want [Y] got [", dst(3), "]"
    stop 1
end if
end program t
