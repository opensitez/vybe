! vybe-test: fortran/keyword_forms/fused_program_unit_keywords
! origin: languages/fortran/tests/fortran/test_keyword_forms.rs

module m
    implicit none
    public

    interface
        subroutine noop()
        end subroutine noop
    end interface

contains
    pure function id(x) result(v)
        integer, intent(in) :: x
        integer :: v
        v = x
    end function id
end module m
