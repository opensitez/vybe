! vybe-test: fortran/procedure_attributes/compile_bind_c_fortran_function
! origin: languages/fortran/tests/fortran/test_procedure_attributes.rs

module c_proc
    use iso_c_binding
    implicit none
contains
    function add_c(a, b) bind(c, name='add_c') result(r)
        integer(c_int), intent(in), value :: a, b
        integer(c_int) :: r
        r = a + b
    end function add_c
end module c_proc

program t
    use c_proc
    print *, "ok"
end program t
