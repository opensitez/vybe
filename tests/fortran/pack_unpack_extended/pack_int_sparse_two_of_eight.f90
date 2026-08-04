! vybe-test: fortran/pack_unpack_extended/pack_int_sparse_two_of_eight
! origin: languages/fortran/tests/fortran/test_pack_unpack_extended.rs
program t
integer :: a(8) = [0, 0, 5, 0, 0, 0, 9, 0]
logical :: mask(8) = [.false., .false., .true., .false., .false., .false., .true., .false.]
integer :: b(2)
b = pack(a, mask)
if ((b(1)) /= 5) then
    print *, "FAIL: want [5] got [", b(1), "]"
    stop 1
end if
if ((b(2)) /= 9) then
    print *, "FAIL: want [9] got [", b(2), "]"
    stop 1
end if
end program t
