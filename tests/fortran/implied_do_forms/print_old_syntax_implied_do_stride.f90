! vybe-test: fortran/implied_do_forms/print_old_syntax_implied_do_stride
! origin: languages/fortran/tests/fortran/test_implied_do_forms.rs
program t
if (trim((/ (i, i = 2, 8, 3) /)) /= "2,5,8") then
    print *, "FAIL: want [2,5,8] got [", (/ (i, i = 2, 8, 3) /), "]"
    stop 1
end if
end program t
