! vybe-test: fortran/keyword_forms/interface_and_binding_keyword_forms
! origin: languages/fortran/tests/fortran/test_keyword_forms.rs

module kw_bind_mod
    use, intrinsic :: iso_c_binding, only: c_int
    implicit none
    contains

subroutine kw_binding(x) bind(c, name="kw_binding")
    integer(c_int), intent(inout) :: x
    x = x
end subroutine kw_binding

subroutine call_kw_binding()
    integer(c_int) :: v
    call kw_binding(v)
end subroutine call_kw_binding
end module kw_bind_mod
