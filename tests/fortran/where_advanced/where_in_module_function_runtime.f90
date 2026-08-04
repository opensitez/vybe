! vybe-test: fortran/where_advanced/where_in_module_function_runtime
! origin: languages/fortran/tests/fortran/test_where_advanced.rs

module where_mod
    implicit none
contains
    function positive_part(a) result(b)
        real, intent(in) :: a(:)
        real :: b(size(a))
        b = 0.0
        where (a > 0.0)
            b = a
        end where
    end function positive_part
end module where_mod

program test
    use where_mod
    real :: v(5) = [-1., 2., -3., 4., -5.]
    real :: p(5)
    p = positive_part(v)
    if ((nint(p(2))) /= 2) then
    print *, "FAIL: want [2] got [", nint(p(2)), "]"
    stop 1
end if
    if ((nint(p(1))) /= 0) then
    print *, "FAIL: want [0] got [", nint(p(1)), "]"
    stop 1
end if
    if ((nint(p(4))) /= 4) then
    print *, "FAIL: want [4] got [", nint(p(4)), "]"
    stop 1
end if
end program test
