! vybe-test: fortran/pack_unpack_extended/pack_int_singleton_selected
! origin: languages/fortran/tests/fortran/test_pack_unpack_extended.rs
program t
integer :: a(5) = [7, 8, 9, 10, 11]
logical :: mask(5) = [.false., .false., .true., .false., .false.]
integer :: b(1)
b = pack(a, mask)
if ((b(1)) /= 9) then
    print *, "FAIL: want [9] got [", b(1), "]"
    stop 1
end if
end program t
