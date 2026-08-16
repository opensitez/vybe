! vybe-test: fortran/strings_extended/verify_all_in_set
! origin: languages/fortran/tests/fortran/test_strings_extended.rs
program t
character(len=10) :: s = 'aabbcc'
if ((verify(s, 'abc')) /= 7) then
    print *, "FAIL: want [7] got [", verify(s, 'abc'), "]"
    stop 1
end if
end program t
