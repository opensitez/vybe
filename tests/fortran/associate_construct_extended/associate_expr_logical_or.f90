! vybe-test: fortran/associate_construct_extended/associate_expr_logical_or
! origin: languages/fortran/tests/fortran/test_associate_construct_extended.rs
program t
logical :: p = .true., q = .false.
associate (either => p .or. q)
if ((either) .neqv. .true.) then
    print *, "FAIL: want [true] got [", either, "]"
    stop 1
end if
end associate
end program t
