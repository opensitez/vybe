! vybe-test: fortran/pack_unpack_extended/unpack_int_fill_from_variable
! origin: languages/fortran/tests/fortran/test_pack_unpack_extended.rs
program t
integer :: a(1) = [99]
logical :: mask(3) = [.false., .true., .false.]
integer :: fill(3) = [1, 2, 3]
integer :: b(3)
b = unpack(a, mask, fill)
if ((b(2)) /= 99) then
    print *, "FAIL: want [99] got [", b(2), "]"
    stop 1
end if
end program t
