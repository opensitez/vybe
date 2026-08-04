! vybe-test: fortran/pack_unpack_extended/pack_char_vowels_from_word
! origin: languages/fortran/tests/fortran/test_pack_unpack_extended.rs
program t
character(len=1) :: a(5) = ['H', 'E', 'L', 'L', 'O']
logical :: mask(5) = [.false., .true., .false., .false., .true.]
character(len=1) :: b(2)
b = pack(a, mask)
if (trim(b(1)) /= "E") then
    print *, "FAIL: want [E] got [", b(1), "]"
    stop 1
end if
if (trim(b(2)) /= "O") then
    print *, "FAIL: want [O] got [", b(2), "]"
    stop 1
end if
end program t
