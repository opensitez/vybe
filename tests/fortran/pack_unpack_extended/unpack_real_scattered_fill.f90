! vybe-test: fortran/pack_unpack_extended/unpack_real_scattered_fill
! origin: languages/fortran/tests/fortran/test_pack_unpack_extended.rs
program t
real :: a(2) = [1.25, 3.75]
logical :: mask(4) = [.true., .false., .true., .false.]
real :: fill(4) = [0.0, 0.0, 0.0, 0.0]
real :: b(4)
b = unpack(a, mask, fill)
if ((int(b(1) * 100)) /= 125) then
    print *, "FAIL: want [125] got [", int(b(1) * 100), "]"
    stop 1
end if
if ((int(b(3) * 100)) /= 375) then
    print *, "FAIL: want [375] got [", int(b(3) * 100), "]"
    stop 1
end if
end program t
