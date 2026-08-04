! vybe-test: fortran/real_ieee_components/rrspacing_of_large_value
! origin: languages/fortran/tests/fortran/test_real_ieee_components.rs
program t
print *, rrspacing(huge(1.0) / 2.0)
end program t
