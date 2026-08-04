! vybe-test: fortran/pure_elemental/contains_multiple_funcs
! origin: languages/fortran/tests/fortran/test_pure_elemental.rs

program test
    print *, add(3, 4)
    print *, mul(3, 4)
contains
    function add(a, b) result(r)
        integer, intent(in) :: a, b
        integer :: r
        r = a + b
    end function add

    function mul(a, b) result(r)
        integer, intent(in) :: a, b
        integer :: r
        r = a * b
    end function mul
end program test
