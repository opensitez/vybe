! vybe-test: fortran/strings_extended/len_padded
! origin: languages/fortran/tests/fortran/test_strings_extended.rs
program t
character(len=10) :: s = 'hi'
if ((len(s)) /= 10) then
    print *, "FAIL: want [10] got [", len(s), "]"
    stop 1
end if
end program t
