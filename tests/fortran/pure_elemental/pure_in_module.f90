! vybe-test: fortran/pure_elemental/pure_in_module
! origin: languages/fortran/tests/fortran/test_pure_elemental.rs

module math_pure
    implicit none
contains
    pure function add(a, b) result(res)
        integer, intent(in) :: a, b
        integer :: res
        res = a + b
    end function add
end module math_pure

program test
    use math_pure
    print *, add(3, 4)
end program test
