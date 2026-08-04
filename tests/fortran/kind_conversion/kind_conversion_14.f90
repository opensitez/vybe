! vybe-test: fortran/kind_conversion/kind_conversion_14
! origin: languages/fortran/tests/fortran/test_kind_conversion.rs
program p
character(len=1) :: c
integer :: i
real :: r
c = 'A'
i = iachar(c)
r = real(i)
print *, i
print *, r
end program p
