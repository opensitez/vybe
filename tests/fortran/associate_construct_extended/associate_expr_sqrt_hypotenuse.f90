! vybe-test: fortran/associate_construct_extended/associate_expr_sqrt_hypotenuse
! origin: languages/fortran/tests/fortran/test_associate_construct_extended.rs
program t
real :: x = 3.0, y = 4.0
associate (hyp => sqrt(x*x + y*y))
if ((int(hyp)) /= 5) then
    print *, "FAIL: want [5] got [", int(hyp), "]"
    stop 1
end if
end associate
end program t
