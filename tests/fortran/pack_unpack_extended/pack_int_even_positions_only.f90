! vybe-test: fortran/pack_unpack_extended/pack_int_even_positions_only
! origin: languages/fortran/tests/fortran/test_pack_unpack_extended.rs
program t
integer :: a(6) = [11, 22, 33, 44, 55, 66]
logical :: mask(6) = [.false., .true., .false., .true., .false., .true.]
integer :: b(3)
b = pack(a, mask)
if ((sum(b)) /= 132) then
    print *, "FAIL: want [132] got [", sum(b), "]"
    stop 1
end if
end program t
