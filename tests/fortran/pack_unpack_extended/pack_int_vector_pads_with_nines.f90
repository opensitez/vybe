! vybe-test: fortran/pack_unpack_extended/pack_int_vector_pads_with_nines
! origin: languages/fortran/tests/fortran/test_pack_unpack_extended.rs
program t
integer :: a(5) = [1, 2, 3, 4, 5]
logical :: mask(5) = [.true., .false., .true., .false., .true.]
integer :: vec(5) = [9, 9, 9, 9, 9]
integer :: b(5)
b = pack(a, mask, vec)
if ((b(1)) /= 1) then
    print *, "FAIL: want [1] got [", b(1), "]"
    stop 1
end if
if ((b(3)) /= 5) then
    print *, "FAIL: want [5] got [", b(3), "]"
    stop 1
end if
if ((b(4)) /= 9) then
    print *, "FAIL: want [9] got [", b(4), "]"
    stop 1
end if
if ((b(5)) /= 9) then
    print *, "FAIL: want [9] got [", b(5), "]"
    stop 1
end if
end program t
