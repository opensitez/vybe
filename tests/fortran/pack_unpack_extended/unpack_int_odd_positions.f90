! vybe-test: fortran/pack_unpack_extended/unpack_int_odd_positions
! origin: languages/fortran/tests/fortran/test_pack_unpack_extended.rs
program t
integer :: a(3) = [5, 15, 25]
logical :: mask(6) = [.true., .false., .true., .false., .true., .false.]
integer :: fill(6) = [0, 0, 0, 0, 0, 0]
integer :: b(6)
b = unpack(a, mask, fill)
if ((b(1)) /= 5) then
    print *, "FAIL: want [5] got [", b(1), "]"
    stop 1
end if
if ((b(3)) /= 15) then
    print *, "FAIL: want [15] got [", b(3), "]"
    stop 1
end if
if ((b(5)) /= 25) then
    print *, "FAIL: want [25] got [", b(5), "]"
    stop 1
end if
end program t
