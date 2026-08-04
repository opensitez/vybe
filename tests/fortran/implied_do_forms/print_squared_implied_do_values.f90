! vybe-test: fortran/implied_do_forms/print_squared_implied_do_values
! origin: languages/fortran/tests/fortran/test_implied_do_forms.rs
program t
if (trim([(i * i, i = 1, 4)]) /= "1,4,9,16") then
    print *, "FAIL: want [1,4,9,16] got [", [(i * i, i = 1, 4)], "]"
    stop 1
end if
end program t
