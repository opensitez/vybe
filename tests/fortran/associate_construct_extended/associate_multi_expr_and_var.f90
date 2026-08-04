! vybe-test: fortran/associate_construct_extended/associate_multi_expr_and_var
! origin: languages/fortran/tests/fortran/test_associate_construct_extended.rs
program t
integer :: n = 5
associate (double => n * 2, orig => n)
if ((double + orig) /= 15) then
    print *, "FAIL: want [15] got [", double + orig, "]"
    stop 1
end if
end associate
end program t
