! vybe-test: fortran/associate_construct_extended/associate_logical_not_expr
! origin: languages/fortran/tests/fortran/test_associate_construct_extended.rs
program t
logical :: p = .false.
associate (q => .not. p)
if ((q) .neqv. .true.) then
    print *, "FAIL: want [true] got [", q, "]"
    stop 1
end if
end associate
end program t
