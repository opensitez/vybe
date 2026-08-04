! vybe-test: fortran/associate_construct_extended/associate_expr_negation
! origin: languages/fortran/tests/fortran/test_associate_construct_extended.rs
program t
integer :: val = 9
associate (neg => -val)
if ((neg) /= -9) then
    print *, "FAIL: want [-9] got [", neg, "]"
    stop 1
end if
end associate
end program t
