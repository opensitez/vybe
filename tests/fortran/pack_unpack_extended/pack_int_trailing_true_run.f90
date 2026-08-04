! vybe-test: fortran/pack_unpack_extended/pack_int_trailing_true_run
! origin: languages/fortran/tests/fortran/test_pack_unpack_extended.rs
program t
integer :: a(6) = [1, 2, 3, 4, 5, 6]
logical :: mask(6) = [.false., .false., .false., .true., .true., .true.]
integer :: b(3)
b = pack(a, mask)
if ((b(1)) /= 4) then
    print *, "FAIL: want [4] got [", b(1), "]"
    stop 1
end if
if ((b(3)) /= 6) then
    print *, "FAIL: want [6] got [", b(3), "]"
    stop 1
end if
end program t
