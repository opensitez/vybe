! vybe-test: fortran/pack_unpack_extended/pack_real_negative_values
! origin: languages/fortran/tests/fortran/test_pack_unpack_extended.rs
program t
real :: a(4) = [-1.0, 2.0, -3.0, 4.0]
logical :: mask(4) = [.true., .false., .true., .false.]
real :: b(2)
b = pack(a, mask)
if ((int(b(1))) /= -1) then
    print *, "FAIL: want [-1] got [", int(b(1)), "]"
    stop 1
end if
if ((int(b(2))) /= -3) then
    print *, "FAIL: want [-3] got [", int(b(2)), "]"
    stop 1
end if
end program t
