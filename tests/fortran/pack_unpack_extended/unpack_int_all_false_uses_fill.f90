! vybe-test: fortran/pack_unpack_extended/unpack_int_all_false_uses_fill
! origin: languages/fortran/tests/fortran/test_pack_unpack_extended.rs
program t
integer :: a(1) = [42]
logical :: mask(3) = [.false., .false., .false.]
integer :: fill(3) = [5, 6, 7]
integer :: b(3)
b = unpack(a, mask, fill)
if ((b(1)) /= 5) then
    print *, "FAIL: want [5] got [", b(1), "]"
    stop 1
end if
if ((b(2)) /= 6) then
    print *, "FAIL: want [6] got [", b(2), "]"
    stop 1
end if
if ((b(3)) /= 7) then
    print *, "FAIL: want [7] got [", b(3), "]"
    stop 1
end if
end program t
