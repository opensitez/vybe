! vybe-test: fortran/strings_extended/index_not_found
! origin: languages/fortran/tests/fortran/test_strings_extended.rs
program t
character(len=20) :: s = 'hello world'
if ((index(s, 'xyz')) /= 0) then
    print *, "FAIL: want [0] got [", index(s, 'xyz'), "]"
    stop 1
end if
end program t
