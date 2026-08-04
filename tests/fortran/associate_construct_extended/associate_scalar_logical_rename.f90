! vybe-test: fortran/associate_construct_extended/associate_scalar_logical_rename
! origin: languages/fortran/tests/fortran/test_associate_construct_extended.rs
program t
logical :: flag = .true.
associate (f => flag)
if ((f) .neqv. .true.) then
    print *, "FAIL: want [true] got [", f, "]"
    stop 1
end if
end associate
end program t
