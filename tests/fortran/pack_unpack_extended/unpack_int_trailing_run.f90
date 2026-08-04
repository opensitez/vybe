! vybe-test: fortran/pack_unpack_extended/unpack_int_trailing_run
! origin: languages/fortran/tests/fortran/test_pack_unpack_extended.rs
program t
integer :: a(2) = [3, 4]
logical :: mask(4) = [.false., .false., .true., .true.]
integer :: fill(4) = [1, 1, 1, 1]
integer :: b(4)
b = unpack(a, mask, fill)
if ((b(1)) /= 1) then
    print *, "FAIL: want [1] got [", b(1), "]"
    stop 1
end if
if ((b(3)) /= 3) then
    print *, "FAIL: want [3] got [", b(3), "]"
    stop 1
end if
if ((b(4)) /= 4) then
    print *, "FAIL: want [4] got [", b(4), "]"
    stop 1
end if
end program t
