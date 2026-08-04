! vybe-test: fortran/interfaces/if_bind_c_optional_42
! origin: languages/fortran/tests/fortran/test_interfaces.rs
subroutine s(x) bind(c, name='s_bind')
integer, intent(in) :: x
end subroutine s
