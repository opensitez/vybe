! vybe-test: fortran/module_use_extended/compile_module_many_contains_procedures
! origin: languages/fortran/tests/fortran/test_module_use_extended.rs

module toolbox
    implicit none
contains
    function f1(x) result(r)
        integer, intent(in) :: x
        integer :: r
        r = x
    end function f1
    function f2(x) result(r)
        integer, intent(in) :: x
        integer :: r
        r = x + 1
    end function f2
    subroutine noop()
    end subroutine noop
end module toolbox

program t
    use toolbox
    print *, f1(1) + f2(2)
end program t
