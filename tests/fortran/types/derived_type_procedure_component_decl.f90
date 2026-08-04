! vybe-test: fortran/types/derived_type_procedure_component_decl
! origin: languages/fortran/tests/fortran/test_types.rs

program test
    implicit none
    abstract interface
        function rhs_func(t) result(v)
            real, intent(in) :: t
            real :: v
        end function rhs_func
    end interface

    type :: CallbackBox
        procedure(rhs_func), pointer, nopass :: fn
    end type CallbackBox

    if (trim("ok") /= "ok") then
    print *, "FAIL: want [ok] got [", "ok", "]"
    stop 1
end if
end program test
