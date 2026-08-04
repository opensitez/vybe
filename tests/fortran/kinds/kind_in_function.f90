! vybe-test: fortran/kinds/kind_in_function
! origin: languages/fortran/tests/fortran/test_kinds.rs

program test
    integer, parameter :: dp = 8
    print *, dp_add(1.0_dp, 2.0_dp)
contains
    function dp_add(a, b) result(res)
        integer, parameter :: dp = 8
        real(kind=dp), intent(in) :: a, b
        real(kind=dp) :: res
        res = a + b
    end function dp_add
end program test
