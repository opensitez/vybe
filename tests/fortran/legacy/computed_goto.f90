! vybe-test: fortran/legacy/computed_goto
! origin: languages/fortran/tests/fortran/test_legacy.rs

program test
    integer :: n = 2
    go to (10, 20, 30), n
10  print *, 'one'; goto 99
20  print *, 'two'; goto 99
30  print *, 'three'
99  continue
end program test
