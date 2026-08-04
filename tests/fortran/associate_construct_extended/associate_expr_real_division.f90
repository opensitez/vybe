! vybe-test: fortran/associate_construct_extended/associate_expr_real_division
! origin: languages/fortran/tests/fortran/test_associate_construct_extended.rs
program t
real :: a = 7.0, b = 2.0
associate (q => a / b)
if ((int(q)) /= 3) then
    print *, "FAIL: want [3] got [", int(q), "]"
    stop 1
end if
end associate
end program t
