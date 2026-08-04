! vybe-test: fortran/kind_conversion/kind_conversion_13
! origin: languages/fortran/tests/fortran/test_kind_conversion.rs
program p
integer, parameter :: k_int = selected_int_kind(2)
real :: r
r = real(3)
print *, int(r, k_int)
end program p
