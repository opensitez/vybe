! vybe-test: fortran/strings_extended/scan_found
! origin: languages/fortran/tests/fortran/test_strings_extended.rs
program t
character(len=10) :: s = 'hello'
if ((scan(s, 'aeiou')) /= 2) then
    print *, "FAIL: want [2] got [", scan(s, 'aeiou'), "]"
    stop 1
end if
end program t
