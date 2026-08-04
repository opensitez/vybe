! vybe-test: fortran/pack_unpack_extended/pack_char_all_true_order
! origin: languages/fortran/tests/fortran/test_pack_unpack_extended.rs
program t
character(len=1) :: a(3) = ['X', 'Y', 'Z']
logical :: mask(3) = [.true., .true., .true.]
character(len=1) :: b(3)
b = pack(a, mask)
if (trim(b(1)) /= "X") then
    print *, "FAIL: want [X] got [", b(1), "]"
    stop 1
end if
if (trim(b(3)) /= "Z") then
    print *, "FAIL: want [Z] got [", b(3), "]"
    stop 1
end if
end program t
