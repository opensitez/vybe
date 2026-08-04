! vybe-test: fortran/pack_unpack_extended/unpack_int_singleton_into_middle
! origin: languages/fortran/tests/fortran/test_pack_unpack_extended.rs
program t
integer :: a(1) = [42]
logical :: mask(5) = [.false., .false., .true., .false., .false.]
integer :: fill(5) = [0, 0, 0, 0, 0]
integer :: b(5)
b = unpack(a, mask, fill)
if ((b(3)) /= 42) then
    print *, "FAIL: want [42] got [", b(3), "]"
    stop 1
end if
end program t
