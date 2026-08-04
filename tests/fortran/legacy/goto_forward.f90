! vybe-test: fortran/legacy/goto_forward
! origin: languages/fortran/tests/fortran/test_legacy.rs

program test
    goto 20
10  print *, 'skip'
    goto 30
20  print *, 'landed'
30  continue
end program test
