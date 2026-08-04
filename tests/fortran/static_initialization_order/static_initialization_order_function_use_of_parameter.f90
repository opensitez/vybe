! vybe-test: fortran/static_initialization_order/static_initialization_order_function_use_of_parameter
! origin: languages/fortran/tests/fortran/test_static_initialization_order.rs

program static_initialization_order_function_use_of_parameter
    integer, parameter :: p = 7
    integer :: q = p
    print *, adjustl(int2str(q))
contains
    character(len=16) function int2str(v)
        integer, intent(in) :: v
        write(int2str, '(I0)') v
    end function int2str
end program static_initialization_order_function_use_of_parameter
