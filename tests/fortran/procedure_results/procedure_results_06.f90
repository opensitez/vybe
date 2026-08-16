! vybe-test: fortran/procedure_results/procedure_results_06
! origin: languages/fortran/tests/fortran/test_procedure_results.rs
program t
if ((f(3)) /= 3) then
    print *, "FAIL: want [3] got [", (f(3)), "]"
    stop 1
end if
contains
integer function f(n)
integer::n
f=n
end function f
end program t
