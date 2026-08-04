! vybe-test: fortran/associate_construct_extended/associate_scalar_integer_rename
! origin: languages/fortran/tests/fortran/test_associate_construct_extended.rs
program t
integer :: n = 42
associate (alias => n)
if ((alias) /= 42) then
    print *, "FAIL: want [42] got [", alias, "]"
    stop 1
end if
end associate
end program t
