! vybe-test: fortran/modules/module_constants
! origin: languages/fortran/tests/fortran/test_modules.rs
module consts
real, parameter :: PI = 3.14159
end module consts
program t
use consts
print *, PI
end program t
