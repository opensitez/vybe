! vybe-test: fortran/pure_elemental/elemental_function_basic
! origin: languages/fortran/tests/fortran/test_pure_elemental.rs

program test
    integer :: a(3) = [1, 2, 3]
    integer :: b(3)
    b = double_elem(a)
    print *, b(1)
contains
    elemental function double_elem(x) result(res)
        integer, intent(in) :: x
        integer :: res
        res = x * 2
    end function double_elem
end program test
