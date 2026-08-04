! vybe-test: fortran/associate_construct_extended/associate_expr_sum_two_vars
! origin: languages/fortran/tests/fortran/test_associate_construct_extended.rs
program t
integer :: a = 10, b = 32
associate (total => a + b)
if ((total) /= 42) then
    print *, "FAIL: want [42] got [", total, "]"
    stop 1
end if
end associate
end program t
