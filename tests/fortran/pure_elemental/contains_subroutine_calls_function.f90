! vybe-test: fortran/pure_elemental/contains_subroutine_calls_function
! origin: languages/fortran/tests/fortran/test_pure_elemental.rs

program test
    call run()
contains
    function compute(x) result(r)
        integer, intent(in) :: x
        integer :: r
        r = x * x + 1
    end function compute

    subroutine run()
        print *, compute(4)
    end subroutine run
end program test
