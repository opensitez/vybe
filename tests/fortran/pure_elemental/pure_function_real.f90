! vybe-test: fortran/pure_elemental/pure_function_real
! origin: languages/fortran/tests/fortran/test_pure_elemental.rs

program test
    print *, hyp(3.0, 4.0)
contains
    pure function hyp(a, b) result(c)
        real, intent(in) :: a, b
        real :: c
        c = sqrt(a*a + b*b)
    end function hyp
end program test
