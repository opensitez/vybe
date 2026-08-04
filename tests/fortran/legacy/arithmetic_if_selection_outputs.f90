! vybe-test: fortran/legacy/arithmetic_if_selection_outputs
! origin: languages/fortran/tests/fortran/test_legacy.rs

program test
    real :: x
    x = -3.0
    if (x) 10, 20, 30
10  print *, -1
20  print *, 0
30  print *, 1
end program test
