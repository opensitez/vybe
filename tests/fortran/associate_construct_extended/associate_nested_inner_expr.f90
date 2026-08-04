! vybe-test: fortran/associate_construct_extended/associate_nested_inner_expr
! origin: languages/fortran/tests/fortran/test_associate_construct_extended.rs
program t
integer :: base = 5
associate (outer => base * 2)
associate (inner => outer + 3)
if ((inner) /= 13) then
    print *, "FAIL: want [13] got [", inner, "]"
    stop 1
end if
end associate
end associate
end program t
