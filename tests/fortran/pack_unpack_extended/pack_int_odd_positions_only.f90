! vybe-test: fortran/pack_unpack_extended/pack_int_odd_positions_only
! origin: languages/fortran/tests/fortran/test_pack_unpack_extended.rs
program t
integer :: a(6) = [11, 22, 33, 44, 55, 66]
logical :: mask(6) = [.true., .false., .true., .false., .true., .false.]
integer :: b(3)
b = pack(a, mask)
if ((sum(b)) /= 99) then
    print *, "FAIL: want [99] got [", sum(b), "]"
    stop 1
end if
end program t
