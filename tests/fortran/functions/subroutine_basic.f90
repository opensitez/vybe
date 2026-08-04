! vybe-test: fortran/functions/subroutine_basic
! origin: languages/fortran/tests/fortran/test_functions.rs

program test
    call greet()
contains
    subroutine greet()
        print *, "Hello"
    end subroutine greet
end program test
