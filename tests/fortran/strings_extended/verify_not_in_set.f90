! vybe-test: fortran/strings_extended/verify_not_in_set
! origin: languages/fortran/tests/fortran/test_strings_extended.rs
program t
character(len=10) :: s = 'hello'
if ((verify(s, 'aeiou')) /= 1) then
    print *, "FAIL: want [1] got [", verify(s, 'aeiou'), "]"
    stop 1
end if
end program t
