! vybe-test: fortran/pack_unpack_extended/pack_real_alternating_halves
! origin: languages/fortran/tests/fortran/test_pack_unpack_extended.rs
program t
real :: a(4) = [1.5, 2.5, 3.5, 4.5]
logical :: mask(4) = [.true., .false., .true., .false.]
real :: b(2)
b = pack(a, mask)
if ((int(b(1) + b(2))) /= 5) then
    print *, "FAIL: want [5] got [", int(b(1) + b(2)), "]"
    stop 1
end if
end program t
