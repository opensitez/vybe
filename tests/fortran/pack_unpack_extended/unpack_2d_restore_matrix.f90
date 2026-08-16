! vybe-test: fortran/pack_unpack_extended/unpack_2d_restore_matrix
! origin: languages/fortran/tests/fortran/test_pack_unpack_extended.rs
program t
integer :: a(3) = [10, 20, 30]
logical :: mask(2,2) = reshape([.true., .false., .true., .false.], [2, 2])
integer :: fill(2,2) = 0
integer :: b(2,2)
b = unpack(a, mask, fill)
if ((b(1,1)) /= 10) then
    print *, "FAIL: want [10] got [", b(1,1), "]"
    stop 1
end if
if ((b(1,2)) /= 20) then
    print *, "FAIL: want [20] got [", b(1,2), "]"
    stop 1
end if
if ((b(2,1)) /= 0) then
    print *, "FAIL: want [0] got [", b(2,1), "]"
    stop 1
end if
end program t
