! vybe-test: fortran/pack_unpack_extended/unpack_char_alternating_fill
! origin: languages/fortran/tests/fortran/test_pack_unpack_extended.rs
program t
character(len=1) :: a(2) = ['M', 'N']
logical :: mask(4) = [.true., .false., .true., .false.]
character(len=1) :: fill(4) = ['.', '.', '.', '.']
character(len=1) :: b(4)
b = unpack(a, mask, fill)
if (trim(b(1)) /= "M") then
    print *, "FAIL: want [M] got [", b(1), "]"
    stop 1
end if
if (trim(b(2)) /= ".") then
    print *, "FAIL: want [.] got [", b(2), "]"
    stop 1
end if
if (trim(b(3)) /= "N") then
    print *, "FAIL: want [N] got [", b(3), "]"
    stop 1
end if
end program t
