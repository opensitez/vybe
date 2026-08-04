! vybe-test: fortran/pack_unpack_extended/pack_2d_row_mask_three_selected
! origin: languages/fortran/tests/fortran/test_pack_unpack_extended.rs
program t
integer :: a(3,2) = reshape([1, 2, 3, 4, 5, 6], [3, 2])
logical :: mask(3,2) = reshape([.true., .true., .false., .false., .true., .false.], [3, 2])
integer :: b(3)
b = pack(a, mask)
if ((sum(b)) /= 9) then
    print *, "FAIL: want [9] got [", sum(b), "]"
    stop 1
end if
end program t
