! vybe-test: fortran/pack_unpack_extended/pack_int_vector_first_last_only
! origin: languages/fortran/tests/fortran/test_pack_unpack_extended.rs
program t
integer :: a(5) = [11, 22, 33, 44, 55]
logical :: mask(5) = [.true., .false., .false., .false., .true.]
integer :: vec(4) = [0, 0, 0, 0]
integer :: b(4)
b = pack(a, mask, vec)
if ((b(1)) /= 11) then
    print *, "FAIL: want [11] got [", b(1), "]"
    stop 1
end if
if ((b(2)) /= 55) then
    print *, "FAIL: want [55] got [", b(2), "]"
    stop 1
end if
if ((b(3)) /= 0) then
    print *, "FAIL: want [0] got [", b(3), "]"
    stop 1
end if
end program t
