! vybe-test: fortran/modules_advanced/module_use_rename
! origin: languages/fortran/tests/fortran/test_modules_advanced.rs

module mathlib
    implicit none
contains
    function square(x) result(r)
        integer, intent(in) :: x
        integer :: r
        r = x * x
    end function square
end module mathlib

program test
    use mathlib, sq => square
    print *, sq(5)
end program test
