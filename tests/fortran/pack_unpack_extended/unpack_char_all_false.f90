! vybe-test: fortran/pack_unpack_extended/unpack_char_all_false
! origin: languages/fortran/tests/fortran/test_pack_unpack_extended.rs
program t
character(len=1) :: a(1) = ['Z']
logical :: mask(3) = [.false., .false., .false.]
character(len=1) :: fill(3) = ['a', 'b', 'c']
character(len=1) :: b(3)
b = unpack(a, mask, fill)
if (trim(b(1)) /= "a") then
    print *, "FAIL: want [a] got [", b(1), "]"
    stop 1
end if
if (trim(b(2)) /= "b") then
    print *, "FAIL: want [b] got [", b(2), "]"
    stop 1
end if
end program t
