! vybe-test: fortran/modules_advanced/interface_explicit
! origin: languages/fortran/tests/fortran/test_modules_advanced.rs

program test
    interface
        function square(x) result(r)
            integer, intent(in) :: x
            integer :: r
        end function square
    end interface
    print *, "ok"
end program test
