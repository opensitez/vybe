! vybe-test: fortran/pack_unpack_extended/pack_int_zeros_in_source
! origin: languages/fortran/tests/fortran/test_pack_unpack_extended.rs
program t
integer :: a(5) = [0, 0, 0, 0, 0]
logical :: mask(5) = [.true., .false., .true., .false., .true.]
integer :: b(3)
b = pack(a, mask)
if ((count(b == 0)) /= 3) then
    print *, "FAIL: want [3] got [", count(b == 0), "]"
    stop 1
end if
end program t
