! vybe-test: fortran/arrays/array_function_result
! origin: languages/fortran/tests/fortran/test_arrays.rs

program test
    integer :: a(3) = [10, 20, 30]
    if ((total(a)) /= 60) then
    print *, "FAIL: want [60] got [", total(a), "]"
    stop 1
end if
contains
    function total(v) result(s)
        integer, intent(in) :: v(:)
        integer :: s
        s = sum(v)
    end function
end program test
