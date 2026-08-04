! vybe-test: fortran/strings_extended/lgt_greater
! origin: languages/fortran/tests/fortran/test_strings_extended.rs
program t
logical :: b
b = lgt('b', 'a')
if ((b) .neqv. .true.) then
    print *, "FAIL: want [true] got [", b, "]"
    stop 1
end if
end program t
