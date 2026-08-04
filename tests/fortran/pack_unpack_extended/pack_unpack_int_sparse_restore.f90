! vybe-test: fortran/pack_unpack_extended/pack_unpack_int_sparse_restore
! origin: languages/fortran/tests/fortran/test_pack_unpack_extended.rs
program t
integer :: src(8) = [0, 0, 5, 0, 0, 0, 9, 0]
logical :: mask(8) = [.false., .false., .true., .false., .false., .false., .true., .false.]
integer :: tmp(2), dst(8)
integer :: fill(8) = [0, 0, 0, 0, 0, 0, 0, 0]
tmp = pack(src, mask)
dst = unpack(tmp, mask, fill)
if ((dst(3)) /= 5) then
    print *, "FAIL: want [5] got [", dst(3), "]"
    stop 1
end if
if ((dst(7)) /= 9) then
    print *, "FAIL: want [9] got [", dst(7), "]"
    stop 1
end if
end program t
