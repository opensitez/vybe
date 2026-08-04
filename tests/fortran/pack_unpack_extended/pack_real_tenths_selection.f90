! vybe-test: fortran/pack_unpack_extended/pack_real_tenths_selection
! origin: languages/fortran/tests/fortran/test_pack_unpack_extended.rs
program t
real :: a(5) = [0.1, 0.2, 0.3, 0.4, 0.5]
logical :: mask(5) = [.false., .true., .false., .true., .false.]
real :: b(2)
b = pack(a, mask)
if ((int(sum(b) * 10)) /= 6) then
    print *, "FAIL: want [6] got [", int(sum(b) * 10), "]"
    stop 1
end if
end program t
