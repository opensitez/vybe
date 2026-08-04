! vybe-test: fortran/pack_unpack_extended/pack_2d_with_vector_pad
! origin: languages/fortran/tests/fortran/test_pack_unpack_extended.rs
program t
integer :: a(2,2) = reshape([1, 2, 3, 4], [2, 2])
logical :: mask(2,2) = reshape([.true., .false., .true., .false.], [2, 2])
integer :: vec(3) = [0, 0, 0]
integer :: b(3)
b = pack(a, mask, vec)
if ((b(1)) /= 1) then
    print *, "FAIL: want [1] got [", b(1), "]"
    stop 1
end if
if ((b(2)) /= 3) then
    print *, "FAIL: want [3] got [", b(2), "]"
    stop 1
end if
if ((b(3)) /= 0) then
    print *, "FAIL: want [0] got [", b(3), "]"
    stop 1
end if
end program t
