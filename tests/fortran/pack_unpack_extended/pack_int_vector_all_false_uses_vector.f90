! vybe-test: fortran/pack_unpack_extended/pack_int_vector_all_false_uses_vector
! origin: languages/fortran/tests/fortran/test_pack_unpack_extended.rs
program t
integer :: a(3) = [1, 2, 3]
logical :: mask(3) = [.false., .false., .false.]
integer :: vec(3) = [88, 77, 66]
integer :: b(3)
b = pack(a, mask, vec)
if ((b(1)) /= 88) then
    print *, "FAIL: want [88] got [", b(1), "]"
    stop 1
end if
if ((b(2)) /= 77) then
    print *, "FAIL: want [77] got [", b(2), "]"
    stop 1
end if
if ((b(3)) /= 66) then
    print *, "FAIL: want [66] got [", b(3), "]"
    stop 1
end if
end program t
