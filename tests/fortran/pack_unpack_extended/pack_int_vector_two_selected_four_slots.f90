! vybe-test: fortran/pack_unpack_extended/pack_int_vector_two_selected_four_slots
! origin: languages/fortran/tests/fortran/test_pack_unpack_extended.rs
program t
integer :: a(4) = [1, 3, 5, 7]
logical :: mask(4) = [.true., .false., .true., .false.]
integer :: vec(4) = [0, 0, 0, 0]
integer :: b(4)
b = pack(a, mask, vec)
if ((b(1)) /= 1) then
    print *, "FAIL: want [1] got [", b(1), "]"
    stop 1
end if
if ((b(2)) /= 5) then
    print *, "FAIL: want [5] got [", b(2), "]"
    stop 1
end if
if ((b(4)) /= 0) then
    print *, "FAIL: want [0] got [", b(4), "]"
    stop 1
end if
end program t
