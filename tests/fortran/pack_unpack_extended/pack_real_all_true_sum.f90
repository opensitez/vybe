! vybe-test: fortran/pack_unpack_extended/pack_real_all_true_sum
! origin: languages/fortran/tests/fortran/test_pack_unpack_extended.rs
program t
real :: a(3) = [0.5, 1.0, 1.5]
logical :: mask(3) = [.true., .true., .true.]
real :: b(3)
b = pack(a, mask)
if ((int(sum(b) * 10)) /= 30) then
    print *, "FAIL: want [30] got [", int(sum(b) * 10), "]"
    stop 1
end if
end program t
