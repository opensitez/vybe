! vybe-test: fortran/subroutine_extended/compile_nested_contains_many_procedures
! origin: languages/fortran/tests/fortran/test_subroutine_extended.rs

program t
    if ((f1(1) + f2(2) + f3(3)) /= 6) then
    print *, "FAIL: want [6] got [", f1(1) + f2(2) + f3(3), "]"
    stop 1
end if
contains
    function f1(x) result(r)
        integer, intent(in) :: x
        integer :: r
        r = x
    end function f1
    function f2(x) result(r)
        integer, intent(in) :: x
        integer :: r
        r = x
    end function f2
    function f3(x) result(r)
        integer, intent(in) :: x
        integer :: r
        r = x
    end function f3
    subroutine noop()
end subroutine noop
end program t
