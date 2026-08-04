! vybe-test: fortran/pack_unpack_extended/pack_char_two_of_four
! origin: languages/fortran/tests/fortran/test_pack_unpack_extended.rs
program t
character(len=1) :: a(4) = ['A', 'B', 'C', 'D']
logical :: mask(4) = [.true., .false., .true., .false.]
character(len=1) :: b(2)
b = pack(a, mask)
if (trim(b(1)) /= "A") then
    print *, "FAIL: want [A] got [", b(1), "]"
    stop 1
end if
if (trim(b(2)) /= "C") then
    print *, "FAIL: want [C] got [", b(2), "]"
    stop 1
end if
end program t
