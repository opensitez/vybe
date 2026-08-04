! vybe-test: fortran/associate_construct_extended/associate_expr_mod_remainder
! origin: languages/fortran/tests/fortran/test_associate_construct_extended.rs
program t
integer :: num = 17, den = 5
associate (rem => mod(num, den))
if ((rem) /= 2) then
    print *, "FAIL: want [2] got [", rem, "]"
    stop 1
end if
end associate
end program t
