! vybe-test: fortran/pack_unpack_extended/pack_unpack_int_roundtrip_first_element
! origin: languages/fortran/tests/fortran/test_pack_unpack_extended.rs
program t
integer :: src(5) = [10, 20, 30, 40, 50]
logical :: mask(5) = [.true., .false., .true., .false., .true.]
integer :: tmp(3), dst(5)
integer :: fill(5) = [-1, -1, -1, -1, -1]
tmp = pack(src, mask)
dst = unpack(tmp, mask, fill)
if ((dst(1)) /= 10) then
    print *, "FAIL: want [10] got [", dst(1), "]"
    stop 1
end if
if ((dst(3)) /= 30) then
    print *, "FAIL: want [30] got [", dst(3), "]"
    stop 1
end if
if ((dst(5)) /= 50) then
    print *, "FAIL: want [50] got [", dst(5), "]"
    stop 1
end if
end program t
