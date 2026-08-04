! vybe-test: fortran/legacy/goto_in_loop
! origin: languages/fortran/tests/fortran/test_legacy.rs

program test
    integer :: i, s
    i = 1
    s = 0
10  if (i > 5) goto 20
    s = s + i
    i = i + 1
    goto 10
20  print *, s
end program test
