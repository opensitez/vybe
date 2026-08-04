! vybe-test: fortran/c_interoperability/c_bind_sub_11
! origin: languages/fortran/tests/fortran/test_c_interoperability.rs
subroutine s() bind(c)
use iso_c_binding
end subroutine s
