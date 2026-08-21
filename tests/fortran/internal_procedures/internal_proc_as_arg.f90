! vybe-test: fortran/internal_procedures/internal_proc_as_arg
! origin: languages/fortran/tests/fortran/test_fortran2008.rs

program test
    integer :: result
    result = apply(3, double_it)
    if ((result) /= 6) then
    print *, "FAIL: want [6] got [", result, "]"
    stop 1
end if
contains
    function apply(x, fn) result(r)
        integer, intent(in) :: x
        interface
            function fn(n) result(v)
                integer, intent(in) :: n
                integer :: v
            end function fn
        end interface
        integer :: r
        r = fn(x)
    end function apply

    function double_it(n) result(v)
        integer, intent(in) :: n
        integer :: v
        v = n * 2
    end function double_it
end program test
