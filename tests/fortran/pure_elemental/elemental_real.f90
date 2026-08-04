! vybe-test: fortran/pure_elemental/elemental_real
! origin: languages/fortran/tests/fortran/test_pure_elemental.rs

program test
    real :: a(4) = [1.0, 4.0, 9.0, 16.0]
    real :: b(4)
    b = root(a)
    print *, b(1)
contains
    elemental function root(x) result(r)
        real, intent(in) :: x
        real :: r
        r = sqrt(x)
    end function root
end program test
