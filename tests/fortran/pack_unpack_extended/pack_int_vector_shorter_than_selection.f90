! vybe-test: fortran/pack_unpack_extended/pack_int_vector_shorter_than_selection
! origin: languages/fortran/tests/fortran/test_pack_unpack_extended.rs
program t
integer :: a(4) = [5, 6, 7, 8]
logical :: mask(4) = [.true., .true., .true., .true.]
integer :: vec(2) = [0, 0]
integer :: b(2)
b = pack(a, mask, vec)
if ((b(1)) /= 5) then
    print *, "FAIL: want [5] got [", b(1), "]"
    stop 1
end if
if ((b(2)) /= 6) then
    print *, "FAIL: want [6] got [", b(2), "]"
    stop 1
end if
end program t
