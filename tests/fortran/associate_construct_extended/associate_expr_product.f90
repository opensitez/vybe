! vybe-test: fortran/associate_construct_extended/associate_expr_product
! origin: languages/fortran/tests/fortran/test_associate_construct_extended.rs
program t
integer :: m = 6, n = 7
associate (prod => m * n)
if ((prod) /= 42) then
    print *, "FAIL: want [42] got [", prod, "]"
    stop 1
end if
end associate
end program t
