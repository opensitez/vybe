! vybe-test: fortran/recursive_results/recursive_results_03
! origin: languages/fortran/tests/fortran/test_recursive_results.rs
recursive complex function f(x) result(r)
complex::x
r=x
end function f
