! vybe-test: fortran/modules_advanced/generic_interface
! origin: languages/fortran/tests/fortran/test_modules_advanced.rs

module generic_mod
    implicit none
    interface my_abs
        module procedure abs_int, abs_real
    end interface
contains
    function abs_int(x) result(r)
        integer, intent(in) :: x
        integer :: r
        r = abs(x)
    end function abs_int
    function abs_real(x) result(r)
        real, intent(in) :: x
        real :: r
        r = abs(x)
    end function abs_real
end module generic_mod

program test
    use generic_mod
    if ((my_abs(-5)) /= 5) then
    print *, "FAIL: want [5] got [", my_abs(-5), "]"
    stop 1
end if
    if ((int(my_abs(-3.14))) /= 3) then
    print *, "FAIL: want [3] got [", int(my_abs(-3.14)), "]"
    stop 1
end if
end program test
