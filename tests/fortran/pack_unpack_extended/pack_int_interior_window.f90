! vybe-test: fortran/pack_unpack_extended/pack_int_interior_window
! origin: languages/fortran/tests/fortran/test_pack_unpack_extended.rs
program t
integer :: a(7) = [100, 2, 3, 4, 5, 6, 200]
logical :: mask(7) = [.false., .true., .true., .true., .true., .true., .false.]
integer :: b(5)
b = pack(a, mask)
if ((b(1)) /= 2) then
    print *, "FAIL: want [2] got [", b(1), "]"
    stop 1
end if
if ((b(5)) /= 6) then
    print *, "FAIL: want [6] got [", b(5), "]"
    stop 1
end if
end program t
