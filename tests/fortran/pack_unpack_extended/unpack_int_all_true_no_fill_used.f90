! vybe-test: fortran/pack_unpack_extended/unpack_int_all_true_no_fill_used
! origin: languages/fortran/tests/fortran/test_pack_unpack_extended.rs
program t
integer :: a(3) = [1, 2, 3]
logical :: mask(3) = [.true., .true., .true.]
integer :: fill(3) = [99, 99, 99]
integer :: b(3)
b = unpack(a, mask, fill)
if ((sum(b)) /= 6) then
    print *, "FAIL: want [6] got [", sum(b), "]"
    stop 1
end if
end program t
