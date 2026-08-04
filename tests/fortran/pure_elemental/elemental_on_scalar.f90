! vybe-test: fortran/pure_elemental/elemental_on_scalar
! origin: languages/fortran/tests/fortran/test_pure_elemental.rs

program test
    print *, cube(3)
contains
    elemental function cube(x) result(res)
        integer, intent(in) :: x
        integer :: res
        res = x * x * x
    end function cube
end program test
