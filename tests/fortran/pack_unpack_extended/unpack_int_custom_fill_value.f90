! vybe-test: fortran/pack_unpack_extended/unpack_int_custom_fill_value
! origin: languages/fortran/tests/fortran/test_pack_unpack_extended.rs
program t
integer :: a(2) = [7, 9]
logical :: mask(4) = [.true., .false., .true., .false.]
integer :: fill(4) = [-1, -1, -1, -1]
integer :: b(4)
b = unpack(a, mask, fill)
if ((b(1)) /= 7) then
    print *, "FAIL: want [7] got [", b(1), "]"
    stop 1
end if
if ((b(2)) /= -1) then
    print *, "FAIL: want [-1] got [", b(2), "]"
    stop 1
end if
if ((b(3)) /= 9) then
    print *, "FAIL: want [9] got [", b(3), "]"
    stop 1
end if
end program t
