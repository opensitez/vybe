! vybe-test: fortran/associate_construct_extended/associate_expr_abs_negative
! origin: languages/fortran/tests/fortran/test_associate_construct_extended.rs
program t
integer :: v = -17
associate (mag => abs(v))
if ((mag) /= 17) then
    print *, "FAIL: want [17] got [", mag, "]"
    stop 1
end if
end associate
end program t
