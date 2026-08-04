! vybe-test: fortran/pack_unpack_extended/pack_int_negative_values_selected
! origin: languages/fortran/tests/fortran/test_pack_unpack_extended.rs
program t
integer :: a(5) = [-3, 2, -7, 4, -1]
logical :: mask(5) = [.true., .false., .true., .false., .true.]
integer :: b(3)
b = pack(a, mask)
if ((b(1)) /= -3) then
    print *, "FAIL: want [-3] got [", b(1), "]"
    stop 1
end if
if ((b(2)) /= -7) then
    print *, "FAIL: want [-7] got [", b(2), "]"
    stop 1
end if
if ((b(3)) /= -1) then
    print *, "FAIL: want [-1] got [", b(3), "]"
    stop 1
end if
end program t
