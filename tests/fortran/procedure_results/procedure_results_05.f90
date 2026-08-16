! vybe-test: fortran/procedure_results/procedure_results_05
! origin: languages/fortran/tests/fortran/test_procedure_results.rs
program t
if (trim(f()) /= "abc") then
    print *, "FAIL: want [abc] got [", trim(f()), "]"
    stop 1
end if
contains
character(len=3) function f()
f='abc'
end function f
end program t
