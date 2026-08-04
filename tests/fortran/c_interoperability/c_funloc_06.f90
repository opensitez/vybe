! vybe-test: fortran/c_interoperability/c_funloc_06
! origin: languages/fortran/tests/fortran/test_c_interoperability.rs
program p
use iso_c_binding
print *, c_associated(c_funloc(s))
contains
subroutine s() bind(c)
end subroutine s
end program p
