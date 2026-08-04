! vybe-test: fortran/pointers/procedure_pointer_swap
! origin: languages/fortran/tests/fortran/test_pointers.rs

program test
    abstract interface
        function unary(x) result(r)
            integer, intent(in) :: x
            integer :: r
        end function unary
    end interface
    procedure(unary), pointer :: fp
    fp => double_it
    print *, fp(3)
    fp => triple_it
    print *, fp(3)
contains
    function double_it(x) result(r)
        integer, intent(in) :: x
        integer :: r
        r = x * 2
    end function double_it
    function triple_it(x) result(r)
        integer, intent(in) :: x
        integer :: r
        r = x * 3
    end function triple_it
end program test
