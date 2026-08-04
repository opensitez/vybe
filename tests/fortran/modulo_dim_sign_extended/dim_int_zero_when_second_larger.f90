! vybe-test: fortran/modulo_dim_sign_extended/dim_int_zero_when_second_larger
! origin: languages/fortran/tests/fortran/test_modulo_dim_sign_extended.rs
program t
if ((dim(3, 10)) /= 0) then
    print *, "FAIL: want [0] got [", dim(3, 10), "]"
    stop 1
end if
end program t
