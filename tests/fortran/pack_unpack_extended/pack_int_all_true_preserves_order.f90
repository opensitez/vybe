! vybe-test: fortran/pack_unpack_extended/pack_int_all_true_preserves_order
! origin: languages/fortran/tests/fortran/test_pack_unpack_extended.rs
program t
integer :: a(4) = [3, 1, 4, 1]
logical :: mask(4) = [.true., .true., .true., .true.]
integer :: b(4)
b = pack(a, mask)
if ((b(2)) /= 1) then
    print *, "FAIL: want [1] got [", b(2), "]"
    stop 1
end if
if ((b(4)) /= 1) then
    print *, "FAIL: want [1] got [", b(4), "]"
    stop 1
end if
end program t
