! vybe-test: fortran/strings_extended/llt_less
! origin: languages/fortran/tests/fortran/test_strings_extended.rs
program t
logical :: b
b = llt('a', 'b')
if ((b) .neqv. .true.) then
    print *, "FAIL: want [true] got [", b, "]"
    stop 1
end if
end program t
