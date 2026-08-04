! vybe-test: fortran/strings_extended/char_slice_from_start
! origin: languages/fortran/tests/fortran/test_strings_extended.rs
program t
character(len=10) :: s = 'hello'
character(len=3) :: sub
sub = s(:3)
if (trim(sub) /= "hel") then
    print *, "FAIL: want [hel] got [", sub, "]"
    stop 1
end if
end program t
