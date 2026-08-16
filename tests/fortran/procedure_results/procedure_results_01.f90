! vybe-test: fortran/procedure_results/procedure_results_01
! origin: languages/fortran/tests/fortran/test_procedure_results.rs
program t
if ((f()) /= 1) then
    print *, "FAIL: want [1] got [", (f()), "]"
    stop 1
end if
contains
integer function f()
f=1
end function f
end program t
