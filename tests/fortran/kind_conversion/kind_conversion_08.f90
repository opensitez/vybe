! vybe-test: fortran/kind_conversion/kind_conversion_08
! origin: languages/fortran/tests/fortran/test_kind_conversion.rs
program p
integer :: i
i = nint(1.6)
print *, i
end program p
