! vybe-test: fortran/associate_construct_extended/associate_expr_logical_and
! origin: languages/fortran/tests/fortran/test_associate_construct_extended.rs
program t
logical :: p = .true., q = .false.
associate (both => p .and. q)
if ((both) .neqv. .false.) then
    print *, "FAIL: want [false] got [", both, "]"
    stop 1
end if
end associate
end program t
