! vybe-test: fortran/pack_unpack_extended/pack_char_with_vector_pad
! origin: languages/fortran/tests/fortran/test_pack_unpack_extended.rs
program t
character(len=1) :: a(2) = ['P', 'Q']
logical :: mask(2) = [.true., .true.]
character(len=1) :: vec(4) = ['-', '-', '-', '-']
character(len=1) :: b(4)
b = pack(a, mask, vec)
if (trim(b(1)) /= "P") then
    print *, "FAIL: want [P] got [", b(1), "]"
    stop 1
end if
if (trim(b(2)) /= "Q") then
    print *, "FAIL: want [Q] got [", b(2), "]"
    stop 1
end if
if (trim(b(3)) /= "-") then
    print *, "FAIL: want [-] got [", b(3), "]"
    stop 1
end if
end program t
