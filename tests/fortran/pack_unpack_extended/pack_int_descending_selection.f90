! vybe-test: fortran/pack_unpack_extended/pack_int_descending_selection
! origin: languages/fortran/tests/fortran/test_pack_unpack_extended.rs
program t
integer :: a(5) = [9, 7, 5, 3, 1]
logical :: mask(5) = [.true., .true., .false., .true., .false.]
integer :: b(3)
b = pack(a, mask)
if ((b(1)) /= 9) then
    print *, "FAIL: want [9] got [", b(1), "]"
    stop 1
end if
if ((b(2)) /= 7) then
    print *, "FAIL: want [7] got [", b(2), "]"
    stop 1
end if
if ((b(3)) /= 3) then
    print *, "FAIL: want [3] got [", b(3), "]"
    stop 1
end if
end program t
