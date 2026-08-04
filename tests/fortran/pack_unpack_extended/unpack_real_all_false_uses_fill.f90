! vybe-test: fortran/pack_unpack_extended/unpack_real_all_false_uses_fill
! origin: languages/fortran/tests/fortran/test_pack_unpack_extended.rs
program t
real :: a(1) = [9.9]
logical :: mask(2) = [.false., .false.]
real :: fill(2) = [1.1, 2.2]
real :: b(2)
b = unpack(a, mask, fill)
if ((int(b(1) * 10)) /= 11) then
    print *, "FAIL: want [11] got [", int(b(1) * 10), "]"
    stop 1
end if
if ((int(b(2) * 10)) /= 22) then
    print *, "FAIL: want [22] got [", int(b(2) * 10), "]"
    stop 1
end if
end program t
