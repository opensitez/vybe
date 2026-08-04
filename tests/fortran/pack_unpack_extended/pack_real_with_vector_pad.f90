! vybe-test: fortran/pack_unpack_extended/pack_real_with_vector_pad
! origin: languages/fortran/tests/fortran/test_pack_unpack_extended.rs
program t
real :: a(3) = [2.0, 4.0, 6.0]
logical :: mask(3) = [.true., .false., .true.]
real :: vec(4) = [0.0, 0.0, 0.0, 0.0]
real :: b(4)
b = pack(a, mask, vec)
if ((int(b(1))) /= 2) then
    print *, "FAIL: want [2] got [", int(b(1)), "]"
    stop 1
end if
if ((int(b(2))) /= 6) then
    print *, "FAIL: want [6] got [", int(b(2)), "]"
    stop 1
end if
if ((int(b(4))) /= 0) then
    print *, "FAIL: want [0] got [", int(b(4)), "]"
    stop 1
end if
end program t
