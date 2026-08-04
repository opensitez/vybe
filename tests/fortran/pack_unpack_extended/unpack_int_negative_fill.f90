! vybe-test: fortran/pack_unpack_extended/unpack_int_negative_fill
! origin: languages/fortran/tests/fortran/test_pack_unpack_extended.rs
program t
integer :: a(2) = [8, 16]
logical :: mask(4) = [.false., .true., .false., .true.]
integer :: fill(4) = [-9, -9, -9, -9]
integer :: b(4)
b = unpack(a, mask, fill)
if ((b(1)) /= -9) then
    print *, "FAIL: want [-9] got [", b(1), "]"
    stop 1
end if
if ((b(2)) /= 8) then
    print *, "FAIL: want [8] got [", b(2), "]"
    stop 1
end if
if ((b(4)) /= 16) then
    print *, "FAIL: want [16] got [", b(4), "]"
    stop 1
end if
end program t
