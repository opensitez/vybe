! vybe-test: fortran/real_ieee_components/fraction_negative_preserves_sign
! origin: languages/fortran/tests/fortran/test_real_ieee_components.rs
program t
print *, fraction(-1.5)
end program t
