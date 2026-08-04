! vybe-test: fortran/stop_error_extended/stop_with_code_expression
! origin: languages/fortran/tests/fortran/test_stop_error_extended.rs
program t
integer :: code = 2
if (code > 1) then
    stop code
else
    print *, 'continue'
end if
print *, 'never'
end program t
