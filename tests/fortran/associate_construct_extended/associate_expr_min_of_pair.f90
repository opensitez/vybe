! vybe-test: fortran/associate_construct_extended/associate_expr_min_of_pair
! origin: languages/fortran/tests/fortran/test_associate_construct_extended.rs
program t
integer :: p = 8, q = 3
associate (lo => min(p, q))
if ((lo) /= 3) then
    print *, "FAIL: want [3] got [", lo, "]"
    stop 1
end if
end associate
end program t
