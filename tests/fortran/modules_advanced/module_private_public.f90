! vybe-test: fortran/modules_advanced/module_private_public
! origin: languages/fortran/tests/fortran/test_modules_advanced.rs

module mymod
    implicit none
    private
    public :: get_value
    integer :: secret = 42
contains
    function get_value() result(v)
        integer :: v
        v = secret
    end function get_value
end module mymod

program test
    use mymod
    print *, get_value()
end program test
