! vybe-test: fortran/strings_extended/scan_not_found
! origin: languages/fortran/tests/fortran/test_strings_extended.rs
program t
character(len=10) :: s = 'bcdfg'
if ((scan(s, 'aeiou')) /= 0) then
    print *, "FAIL: want [0] got [", scan(s, 'aeiou'), "]"
    stop 1
end if
end program t
