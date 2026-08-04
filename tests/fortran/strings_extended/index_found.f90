! vybe-test: fortran/strings_extended/index_found
! origin: languages/fortran/tests/fortran/test_strings_extended.rs
program t
character(len=20) :: s = 'hello world'
if ((index(s, 'world')) /= 7) then
    print *, "FAIL: want [7] got [", index(s, 'world'), "]"
    stop 1
end if
end program t
