! vybe-test: fortran/fortran2018_extended/reduce_builtin_min_operator
! origin: languages/fortran/tests/fortran/test_fortran2018_extended.rs
program t
integer :: a(4) = [3, 1, 4, 2]
if ((reduce(a, operator(min))) /= 1) then
    print *, "FAIL: want [1] got [", reduce(a, operator(min)), "]"
    stop 1
end if
end program t
