! vybe-test: fortran/pack_unpack_extended/pack_unpack_int_count_nonzero
! origin: languages/fortran/tests/fortran/test_pack_unpack_extended.rs
program t
integer :: src(6) = [1, 0, 2, 0, 3, 0]
logical :: mask(6) = [.true., .false., .true., .false., .true., .false.]
integer :: tmp(3), dst(6)
integer :: fill(6) = [0, 0, 0, 0, 0, 0]
tmp = pack(src, mask)
dst = unpack(tmp, mask, fill)
if ((count(dst > 0)) /= 3) then
    print *, "FAIL: want [3] got [", count(dst > 0), "]"
    stop 1
end if
end program t
