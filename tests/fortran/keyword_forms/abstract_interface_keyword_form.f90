! vybe-test: fortran/keyword_forms/abstract_interface_keyword_form
! origin: languages/fortran/tests/fortran/test_keyword_forms.rs

module kw_abstract_iface
interface
    subroutine abstract_target(x)
        integer, intent(inout) :: x
    end subroutine abstract_target
end interface

abstract interface
    subroutine abstract_target(x)
        integer, intent(inout) :: x
    end subroutine abstract_target
end interface
end module kw_abstract_iface
