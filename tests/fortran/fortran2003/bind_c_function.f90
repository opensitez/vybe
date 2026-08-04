! vybe-test: fortran/fortran2003/bind_c_function
! origin: languages/fortran/tests/fortran/test_fortran2003.rs

module c_funcs
    use iso_c_binding
    implicit none
    interface
        function c_strlen(s) bind(c, name='strlen') result(n)
            use iso_c_binding
            type(c_ptr), value :: s
            integer(c_size_t) :: n
        end function c_strlen
    end interface
end module c_funcs

program test
    print *, "ok"
end program test
