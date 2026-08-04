! vybe-test: fortran/pure_elemental/pure_function_basic
! origin: languages/fortran/tests/fortran/test_pure_elemental.rs

program test
    print *, square(5)
contains
    pure function square(x) result(res)
        integer, intent(in) :: x
        integer :: res
        res = x * x
    end function square
end program test
