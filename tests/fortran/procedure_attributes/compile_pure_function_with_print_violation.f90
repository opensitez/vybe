! vybe-test: fortran/procedure_attributes/compile_pure_function_with_print_violation
! origin: languages/fortran/tests/fortran/test_procedure_attributes.rs

program t
    print *, bad_pure(5)
contains
    pure function bad_pure(x) result(r)
        integer, intent(in) :: x
        integer :: r
        print *, x
        r = x * 2
    end function bad_pure
end program t
