! vybe-test: fortran/pointers/procedure_pointer_basic
! origin: languages/fortran/tests/fortran/test_pointers.rs

program test
    procedure(int_fn), pointer :: fp => null()
    fp => double_it
    print *, fp(5)
contains
    function double_it(x) result(r)
        integer, intent(in) :: x
        integer :: r
        r = x * 2
    end function double_it
    function int_fn(x) result(r)
        integer, intent(in) :: x
        integer :: r
    end function int_fn
end program test
