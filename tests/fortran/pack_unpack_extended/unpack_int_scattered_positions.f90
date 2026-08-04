! vybe-test: fortran/pack_unpack_extended/unpack_int_scattered_positions
! origin: languages/fortran/tests/fortran/test_pack_unpack_extended.rs
program t
integer :: a(3) = [2, 4, 6]
logical :: mask(5) = [.true., .false., .true., .false., .true.]
integer :: fill(5) = [0, 0, 0, 0, 0]
integer :: b(5)
b = unpack(a, mask, fill)
if ((b(1)) /= 2) then
    print *, "FAIL: want [2] got [", b(1), "]"
    stop 1
end if
if ((b(3)) /= 4) then
    print *, "FAIL: want [4] got [", b(3), "]"
    stop 1
end if
if ((b(5)) /= 6) then
    print *, "FAIL: want [6] got [", b(5), "]"
    stop 1
end if
end program t
