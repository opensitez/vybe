! vybe-test: fortran/associate_construct_extended/associate_expr_max_of_pair
! origin: languages/fortran/tests/fortran/test_associate_construct_extended.rs
program t
integer :: p = 8, q = 3
associate (hi => max(p, q))
if ((hi) /= 8) then
    print *, "FAIL: want [8] got [", hi, "]"
    stop 1
end if
end associate
end program t
