! vybe-test: fortran/strings_extended/char_slice_to_end
! origin: languages/fortran/tests/fortran/test_strings_extended.rs
program t
character(len=10) :: s = 'hello'
character(len=5) :: sub
sub = s(3:)
if (trim(trim(sub)) /= "llo") then
    print *, "FAIL: want [llo] got [", trim(sub), "]"
    stop 1
end if
end program t
