! vybe-test: fortran/procedure_associated_pointer_calls/procedure_associated_pointer_calls_nesting_pointer_to_component
! origin: languages/fortran/tests/fortran/test_procedure_associated_pointer_calls.rs

program procedure_associated_pointer_calls_nesting_pointer_to_component
    type callback_holder
        procedure(scale_iface), pointer, nopass :: f
    end type callback_holder

    abstract interface
        integer function scale_iface(x, factor)
            integer, intent(in) :: x
            integer, intent(in), optional :: factor
        end function scale_iface
    end interface

    type(callback_holder) :: holder
    integer :: value

    holder%f => scale
    value = holder%f(4, 3)
    if ((value) /= 12) then
    print *, "FAIL: want [12] got [", value, "]"
    stop 1
end if
contains
    integer function scale(x, factor)
        integer, intent(in) :: x
        integer, intent(in), optional :: factor
        if (present(factor)) then
            scale = x * factor
        else
            scale = x
        end if
    end function scale
end program procedure_associated_pointer_calls_nesting_pointer_to_component
