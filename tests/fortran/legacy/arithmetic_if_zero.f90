! vybe-test: fortran/legacy/arithmetic_if_zero
! origin: languages/fortran/tests/fortran/test_legacy.rs

program test
    real :: x = 0.0
    if (x) 10, 20, 30
10  print *, 'negative'; goto 99
20  print *, 'zero'; goto 99
30  print *, 'positive'
99  continue
end program test
