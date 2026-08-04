! vybe-test: fortran/strings_extended/char_slice_range
! origin: languages/fortran/tests/fortran/test_strings_extended.rs
program t
character(len=10) :: s = 'abcdefgh'
character(len=3) :: sub
sub = s(2:4)
if (trim(sub) /= "bcd") then
    print *, "FAIL: want [bcd] got [", sub, "]"
    stop 1
end if
end program t
