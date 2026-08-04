! vybe-test: fortran/recursive_results/recursive_results_06
! origin: languages/fortran/tests/fortran/test_recursive_results.rs
recursive integer function f(n) result(r)
integer::n
r=n
end function f
