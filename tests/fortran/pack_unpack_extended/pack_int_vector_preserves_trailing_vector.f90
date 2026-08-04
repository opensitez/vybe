! vybe-test: fortran/pack_unpack_extended/pack_int_vector_preserves_trailing_vector
! origin: languages/fortran/tests/fortran/test_pack_unpack_extended.rs
program t
integer :: a(5) = [3, 6, 9, 12, 15]
logical :: mask(5) = [.false., .true., .false., .true., .false.]
integer :: vec(4) = [100, 200, 300, 400]
integer :: b(4)
b = pack(a, mask, vec)
if ((b(1)) /= 6) then
    print *, "FAIL: want [6] got [", b(1), "]"
    stop 1
end if
if ((b(2)) /= 12) then
    print *, "FAIL: want [12] got [", b(2), "]"
    stop 1
end if
if ((b(3)) /= 300) then
    print *, "FAIL: want [300] got [", b(3), "]"
    stop 1
end if
end program t
