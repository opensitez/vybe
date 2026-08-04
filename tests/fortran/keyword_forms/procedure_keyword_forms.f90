! vybe-test: fortran/keyword_forms/procedure_keyword_forms
! origin: languages/fortran/tests/fortran/test_keyword_forms.rs

module kw_proc
    implicit none
contains
    recursive integer function fact(n) result(r)
        integer, intent(in) :: n
        if (n <= 1) then
            r = 1
        else
            r = n * fact(n - 1)
        end if
    end function fact

    pure elemental integer function inc(x) result(r)
        integer, intent(in) :: x
        r = x + 1
    end function inc

    module subroutine noop_sub(x)
        integer, intent(inout) :: x
        x = x
    end subroutine noop_sub
end module kw_proc
