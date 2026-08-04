! vybe-test: fortran/pack_unpack_extended/pack_int_alternating_mask_first_third_fifth
! origin: languages/fortran/tests/fortran/test_pack_unpack_extended.rs
program t
integer :: a(5) = [10, 20, 30, 40, 50]
logical :: mask(5) = [.true., .false., .true., .false., .true.]
integer :: b(3)
b = pack(a, mask)
if ((b(1)) /= 10) then
    print *, "FAIL: want [10] got [", b(1), "]"
    stop 1
end if
if ((b(2)) /= 30) then
    print *, "FAIL: want [30] got [", b(2), "]"
    stop 1
end if
if ((b(3)) /= 50) then
    print *, "FAIL: want [50] got [", b(3), "]"
    stop 1
end if
end program t
