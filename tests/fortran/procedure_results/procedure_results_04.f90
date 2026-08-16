! vybe-test: fortran/procedure_results/procedure_results_04
! origin: languages/fortran/tests/fortran/test_procedure_results.rs
program t
if (merge(1, 0, f()) /= 1) then
    print *, "FAIL: want [1] got [", merge(1, 0, f()), "]"
    stop 1
end if
contains
logical function f()
f=.true.
end function f
end program t
