! vybe-test: fortran/pure_elemental/intent_in_multiple
! origin: languages/fortran/tests/fortran/test_pure_elemental.rs

program test
    print *, add3(1, 2, 3)
contains
    function add3(a, b, c) result(res)
        integer, intent(in) :: a, b, c
        integer :: res
        res = a + b + c
    end function add3
end program test
