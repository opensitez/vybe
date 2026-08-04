! vybe-test: fortran/kind_conversion/kind_conversion_01
! origin: languages/fortran/tests/fortran/test_kind_conversion.rs
program p
integer :: i
real :: r=1.5
i = int(r)
print *, i
end program p
