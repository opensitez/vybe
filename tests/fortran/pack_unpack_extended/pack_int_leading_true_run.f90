! vybe-test: fortran/pack_unpack_extended/pack_int_leading_true_run
! origin: languages/fortran/tests/fortran/test_pack_unpack_extended.rs
program t
integer :: a(6) = [1, 2, 3, 4, 5, 6]
logical :: mask(6) = [.true., .true., .true., .false., .false., .false.]
integer :: b(3)
b = pack(a, mask)
if ((sum(b)) /= 6) then
    print *, "FAIL: want [6] got [", sum(b), "]"
    stop 1
end if
end program t
