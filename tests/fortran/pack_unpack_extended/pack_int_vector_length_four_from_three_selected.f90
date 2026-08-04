! vybe-test: fortran/pack_unpack_extended/pack_int_vector_length_four_from_three_selected
! origin: languages/fortran/tests/fortran/test_pack_unpack_extended.rs
program t
integer :: a(6) = [2, 4, 6, 8, 10, 12]
logical :: mask(6) = [.true., .true., .true., .false., .false., .false.]
integer :: vec(4) = [-1, -1, -1, -1]
integer :: b(4)
b = pack(a, mask, vec)
if ((b(1)) /= 2) then
    print *, "FAIL: want [2] got [", b(1), "]"
    stop 1
end if
if ((b(3)) /= 6) then
    print *, "FAIL: want [6] got [", b(3), "]"
    stop 1
end if
if ((b(4)) /= -1) then
    print *, "FAIL: want [-1] got [", b(4), "]"
    stop 1
end if
end program t
