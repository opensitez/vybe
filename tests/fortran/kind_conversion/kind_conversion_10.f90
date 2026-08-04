! vybe-test: fortran/kind_conversion/kind_conversion_10
! origin: languages/fortran/tests/fortran/test_kind_conversion.rs
program p
integer :: i
real :: r=1.0
i = transfer(r, i)
print *, i
end program p
