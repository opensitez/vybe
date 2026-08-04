! vybe-test: fortran/pack_unpack_extended/unpack_int_leading_run
! origin: languages/fortran/tests/fortran/test_pack_unpack_extended.rs
program t
integer :: a(2) = [100, 200]
logical :: mask(4) = [.true., .true., .false., .false.]
integer :: fill(4) = [0, 0, 0, 0]
integer :: b(4)
b = unpack(a, mask, fill)
if ((b(1)) /= 100) then
    print *, "FAIL: want [100] got [", b(1), "]"
    stop 1
end if
if ((b(2)) /= 200) then
    print *, "FAIL: want [200] got [", b(2), "]"
    stop 1
end if
if ((b(4)) /= 0) then
    print *, "FAIL: want [0] got [", b(4), "]"
    stop 1
end if
end program t
