! vybe-test: fortran/modules/use_only
! origin: languages/fortran/tests/fortran/test_modules.rs
module mymod
integer :: a = 10
integer :: b = 20
end module mymod
program t
use mymod, only: a
print *, a
end program t
