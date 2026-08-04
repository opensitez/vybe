! vybe-test: fortran/real_ieee_components/spacing_at_tiny_value
! origin: languages/fortran/tests/fortran/test_real_ieee_components.rs
program t
print *, spacing(tiny(1.0))
end program t
