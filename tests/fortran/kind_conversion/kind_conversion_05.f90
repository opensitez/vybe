! vybe-test: fortran/kind_conversion/kind_conversion_05
! origin: languages/fortran/tests/fortran/test_kind_conversion.rs
program p
integer(kind=8) :: i
i = int(1.5, kind=8)
print *, i
end program p
