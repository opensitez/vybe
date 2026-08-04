! vybe-test: fortran/strings_extended/len_trim_padded
! origin: languages/fortran/tests/fortran/test_strings_extended.rs
program t
character(len=10) :: s = 'hi'
if ((len_trim(s)) /= 2) then
    print *, "FAIL: want [2] got [", len_trim(s), "]"
    stop 1
end if
end program t
