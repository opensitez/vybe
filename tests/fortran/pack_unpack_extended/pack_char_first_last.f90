! vybe-test: fortran/pack_unpack_extended/pack_char_first_last
! origin: languages/fortran/tests/fortran/test_pack_unpack_extended.rs
program t
character(len=1) :: a(5) = ['1', '2', '3', '4', '5']
logical :: mask(5) = [.true., .false., .false., .false., .true.]
character(len=1) :: b(2)
b = pack(a, mask)
if ((b(1)) /= '1') then
    print *, "FAIL: want [1] got [", b(1), "]"
    stop 1
end if
if ((b(2)) /= '5') then
    print *, "FAIL: want [5] got [", b(2), "]"
    stop 1
end if
end program t
