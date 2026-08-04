! vybe-test: fortran/pack_unpack_extended/pack_int_vector_longer_than_selection
! origin: languages/fortran/tests/fortran/test_pack_unpack_extended.rs
program t
integer :: a(3) = [7, 14, 21]
logical :: mask(3) = [.true., .false., .true.]
integer :: vec(5) = [1, 2, 3, 4, 5]
integer :: b(5)
b = pack(a, mask, vec)
if ((b(1)) /= 7) then
    print *, "FAIL: want [7] got [", b(1), "]"
    stop 1
end if
if ((b(2)) /= 21) then
    print *, "FAIL: want [21] got [", b(2), "]"
    stop 1
end if
if ((b(5)) /= 5) then
    print *, "FAIL: want [5] got [", b(5), "]"
    stop 1
end if
end program t
