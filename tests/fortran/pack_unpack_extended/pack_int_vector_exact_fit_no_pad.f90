! vybe-test: fortran/pack_unpack_extended/pack_int_vector_exact_fit_no_pad
! origin: languages/fortran/tests/fortran/test_pack_unpack_extended.rs
program t
integer :: a(5) = [2, 4, 6, 8, 10]
logical :: mask(5) = [.true., .false., .true., .false., .true.]
integer :: vec(3) = [0, 0, 0]
integer :: b(3)
b = pack(a, mask, vec)
if ((sum(b)) /= 18) then
    print *, "FAIL: want [18] got [", sum(b), "]"
    stop 1
end if
end program t
