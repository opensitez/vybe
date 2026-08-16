! vybe-test: fortran/recursive_results/recursive_results_03
! origin: languages/fortran/tests/fortran/test_recursive_results.rs
program t
if (nint(real(f((1.0,2.0))) + aimag(f((1.0,2.0)))) /= 3) then
    print *, "FAIL: want [3] got [", nint(real(f((1.0,2.0))) + aimag(f((1.0,2.0)))), "]"
    stop 1
end if
contains
recursive complex function f(x) result(r)
complex::x
r=x
end function f
end program t
