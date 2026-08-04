! vybe-test: fortran/pack_unpack_extended/unpack_int_even_positions
! origin: languages/fortran/tests/fortran/test_pack_unpack_extended.rs
program t
integer :: a(3) = [10, 20, 30]
logical :: mask(6) = [.false., .true., .false., .true., .false., .true.]
integer :: fill(6) = [0, 0, 0, 0, 0, 0]
integer :: b(6)
b = unpack(a, mask, fill)
if ((b(2)) /= 10) then
    print *, "FAIL: want [10] got [", b(2), "]"
    stop 1
end if
if ((b(4)) /= 20) then
    print *, "FAIL: want [20] got [", b(4), "]"
    stop 1
end if
if ((b(6)) /= 30) then
    print *, "FAIL: want [30] got [", b(6), "]"
    stop 1
end if
end program t
