! vybe-test: fortran/pack_unpack_extended/pack_2d_by_column_major_mask
! origin: languages/fortran/tests/fortran/test_pack_unpack_extended.rs
program t
integer :: a(2,3) = reshape([1, 4, 2, 5, 3, 6], [2, 3])
logical :: mask(2,3) = reshape([.true., .false., .true., .false., .true., .false.], [2, 3])
integer :: b(3)
b = pack(a, mask)
if ((b(1)) /= 1) then
    print *, "FAIL: want [1] got [", b(1), "]"
    stop 1
end if
if ((b(2)) /= 2) then
    print *, "FAIL: want [2] got [", b(2), "]"
    stop 1
end if
if ((b(3)) /= 3) then
    print *, "FAIL: want [3] got [", b(3), "]"
    stop 1
end if
end program t
